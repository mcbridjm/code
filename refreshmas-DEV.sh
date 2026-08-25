#!/bin/bash
# File will be outputted to the DATA_PUMP_DIR (/u02/export)
#
SCRPTPATH=/u01/scripts/migration
. $SCRPTPATH/master.properties
##read -p "Enter an export filename e.g. MAS_filename_date.dmp: " DMPNAME
#DBNAME='MASSBX'
DATETEMP=`date +%Y%m%d_%H%M%S`
IMPMASTEMPLATE="import_mas_template_$DATETEMP.sql"
ORACLEUSERDIR=/u01/scripts

# Avoids exposing credentials in the process list via ps
run_sqlplus() {
    local user="$1" pass="$2" db="$3" script="$4"
    sqlplus -s /nolog <<EOF
CONNECT $user/$pass@$db
@$script
EXIT
EOF
}

read -p "You must login to OpenShift with your token. Enter it here: " OCTOKEN
OCLOGIN=`oc login --token=sha256~$OCTOKEN --server=https://api.vapo-aws-valip-mxd.va.gov:6443`
oc whoami

read -p "Enter Oracle_SID: [MASSBX,MASDEMO] " DBNAME
if [ -f $SCRPTPATH/$DBNAME-preimport.sql ];
then
        echo "Found Preimport Script, continuing"
        #Set environment variables and pathing for DB
        . $ORACLEUSERDIR/$DBNAME'.env'
        DMPPATH=/u06/export/$DBNAME
else
        echo "$SCRPTPATH/$DBNAME-preimport.sql does not exist, aborting"
        exit
fi

export ORACLE_SID=$DBNAME
echo $DMPNAME $DBNAME $ORACLE_SID
#
#
#
# Update these to match your OC secret name and namespace
OC_CRED_SECRET="mas-db-credentials"
OC_NAMESPACE="mas-${DBNAME,,}-core"

if oc get secret "$OC_CRED_SECRET" -n "$OC_NAMESPACE" &>/dev/null; then
    echo "Fetching credentials from OpenShift secret: $OC_CRED_SECRET"
    VAMASADMIN_PASS=$(oc get secret "$OC_CRED_SECRET" -n "$OC_NAMESPACE" -o jsonpath='{.data.vamasadmin_password}' | base64 -d)
    DB_PASS=$(oc get secret "$OC_CRED_SECRET" -n "$OC_NAMESPACE" -o jsonpath='{.data.maximo_password}' | base64 -d)
else
    echo "OC secret '$OC_CRED_SECRET' not found in '$OC_NAMESPACE' — prompting manually."
    echo "Enter vamasadmin password: "
    read -s VAMASADMIN_PASS
    echo "Enter Maximo schema password: "
    read -s DB_PASS
fi
run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" checkconn.sql
run_sqlplus maximo "$DB_PASS" "$DBNAME" checkconn.sql

ENVNAME=${DBNAME,,}
/u01/scripts/servertools.sh stop $ENVNAME
# read -p "Enter Import Name: " EXPNAME
#
cp $SCRPTPATH/import_mas_template.sql $SCRPTPATH/tmp/$IMPMASTEMPLATE
sed -i "s/DBNAME/$EXPNAME/g" $SCRPTPATH/tmp/$IMPMASTEMPLATE
## Uncomment below double commented lines for password change logic
CHGPW="create user maximo identified by "$DB_PASS";"


#echo $CHGPW
#echo $DB_PASS
eval DB_PASS=$`echo $DBNAME"_PASS"`

##sed -i "2s/.*/$CHGPW/" $SCRPTPATH/$DBNAME-preimport.sql

#read -p "Pausing. Did pre-import run? " temp

# Call the import
# Import script commented out as import requires retooling for RDS
echo "This will wipe the target database's maximo schema. "
echo
while true; do
        read -p "Continue Y/N? " ABORTPRMPT
        case $ABORTPRMPT in
        Y ) break;;
        N ) exit;;
        * ) echo "Enter Y or N ";;
        esac
done
#
# set passwords as needed without leaving it exposed
#
cp $SCRPTPATH/$DBNAME-preimport.sql $SCRPTPATH/tmp/$DBNAME-preimport-$DATETEMP.sql
sed -i "s/PASSWORDVALUE/$DB_PASS/g" $SCRPTPATH/tmp/$DBNAME-preimport-$DATETEMP.sql
#
# Prep the DB
#
echo "Deleting dump files from RDS"
run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" cleanup_dmpfiles.sql
echo "Killing connected users"
run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" kill_connected.sql
echo "Removing schema"
run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" "$SCRPTPATH/tmp/$DBNAME-preimport-$DATETEMP.sql"

while true; do
    read -p "Did the drop succeed?" ABORTPRMPT
    case $ABORTPRMPT in
           Y ) break;;
           N ) run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" "$SCRPTPATH/tmp/$DBNAME-preimport-$DATETEMP.sql" ;;
           * ) echo "Enter Y or N ";;
    esac
done

echo "Run SELECT * FROM table(rdsadmin.rds_file_util.listdir('DATA_PUMP_DIR')) order by mtime; until there are no incomplete A_MASPRD-CURRENT.dmp or B_MASPRD-CURRENT.dmp files listed. Approximately 5 minutes. You can just hit '/' and then 'Enter'"
run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" copy_to_rds.sql
run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" listdpdir.sql

while true; do
    read -p "Did the copy finish?" ABORTPRMPT
    case $ABORTPRMPT in
           Y ) break;;
           N ) run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" listdpdir.sql ;;
           * ) echo "Enter Y or N ";;
    esac
done

#
# Call Import
#read -p "Did the Cleanup run? Hint Enter or ctrl+c " ABORTPRMPT
#
IMPDP_PARFILE=$(mktemp /tmp/impdp_XXXXXX.par)
chmod 600 "$IMPDP_PARFILE"
echo "userid=vamasadmin/$VAMASADMIN_PASS@$DBNAME" > "$IMPDP_PARFILE"

impdp parfile="$IMPDP_PARFILE" directory=DATA_PUMP_DIR dumpfile=A_MASPRD-CURRENT.dmp TABLE_EXISTS_ACTION=REPLACE logfile=A_$DBNAME-import-$DATETEMP.log schemas=MAXIMO
impdp parfile="$IMPDP_PARFILE" directory=DATA_PUMP_DIR dumpfile=B_MASPRD-CURRENT.dmp TABLE_EXISTS_ACTION=REPLACE logfile=B_$DBNAME-import-$DATETEMP.log tables=maximo.ASSET,maximo.A_ASSET,maximo.ASSETTRANS,maximo.WORKORDER,maximo.A_WORKORDER,maximo.A_LONGDESCRIPTION,maximo.LOGINTRACKING

rm -f "$IMPDP_PARFILE"
#read -p "Pausing. Did the import fire?. Go check "
# Cleanup after when we have fired the job
rm $SCRPTPATH/tmp/$DBNAME-preimport-$DATETEMP.sql

#sqlplus maximo/$DB_PASS@$DBNAME @'/'$SCRPTPATH'/dba.sql'
#read -p "Pausing. Did dba.sql run? " temp
run_sqlplus maximo "$DB_PASS" "$DBNAME" "$SCRPTPATH/common.sql"
#read -p "Pausing. Did common.sql run? " temp

# Run the maxiface setup script
run_sqlplus vamasadmin "$VAMASADMIN_PASS" "$DBNAME" "$SCRPTPATH/maxiface-common.sql"
# Define maxiface account
#sqlplus maxiface/$MAXIFACE_PASS@$DBNAME @'/'$SCRPTPATH'/'maxiface-common.sql
#read -p "Pausing. Did maxiface-common run? " temp

# Standard 7.x transforms to rename an environment
run_sqlplus maximo "$DB_PASS" "$DBNAME" "$SCRPTPATH/$DBNAME-postimport.sql"

#read -p "Pausing. Did the post-import run? " temp
#
while true; do
        read -p "Run the Master Cleanse? Y|N " ABORTPRMPT
        case $ABORTPRMPT in
        Y )
                run_sqlplus maximo "$DB_PASS" "$DBNAME" "$SCRPTPATH/cleanse-master.sql"
                break;;
        N )
                break;;
        * ) echo "Enter Y or N ";;
        esac
done
read -p "Pausing. Did cleanse-master run? " temp
# Upgrade specific steps.
ENVNAME=${DBNAME,,}
/u01/scripts/servertools.sh start $ENVNAME
sleep 60
oc project mas-$ENVNAME-manage
sleep 300
MAXINSTPOD=`oc get pods |grep maxinst |grep Running |awk '{print $1}'`
sleep 300
oc exec -it $MAXINSTPOD '/opt/IBM/SMP/maximo/tools/maximo/resetcryptocryptox.sh'
# Note it may be necessary to reimport client certificates if applications refuse to show in the navigation sidebar. To do so, run: ./insertClientCertAPIKeys.sh
oc exec $MAXINSTPOD -- bash /opt/IBM/SMP/maximo/tools/maximo/insertClientCertAPIKeys.sh

#oc exec -it $MAXINSTPOD /opt/IBM/SMP/maximo/tools/maximo/syncmetadata.sh
run_sqlplus maximo "$DB_PASS" "$DBNAME" "$SCRPTPATH/MAS-postupgrade.sql"


#while true; do
#       read -p "Is this a new upgrade? (Select 'Y' for a 7.6 import) Y|N " ABORTPRMPT2
#        case $ABORTPRMPT2 in
#        Y )
#               sqlplus maximo/$DB_PASS@$DBNAME @'/'$SCRPTPATH'/'MAS-upgrade-common.sql
#               read -p "Return here once the upgrade is complete AND MANAGE IS STARTED. Alternatively ctrl+c and run MAS-postupgrade.sql! " ABORTPRMPT3
#               sqlplus maximo/$DB_PASS@$DBNAME @'/'$SCRPTPATH'/MAS-postupgrade.sql'
#
#               #oc patch secret main-manage-encryptionsecret-operator -n mas-$DBNAME-manage --type=merge -f oc-patch-secret.yaml
#               exit;;
#        N ) exit;;
#        * ) echo "Enter Y or N ";;
#        esac
#done
