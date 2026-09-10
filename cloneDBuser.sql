
SET PAGESIZE 0 FEEDBACK OFF VERIFY OFF HEADING OFF

ACCEPT SOURCE_USER CHAR PROMPT 'Enter SOURCE user: '
ACCEPT TARGET_USER CHAR PROMPT 'Enter TARGET user: '

-- ROLE GRANTS --
SELECT 'GRANT ' || granted_role || ' TO ' || UPPER('&&TARGET_USER') ||
       CASE WHEN admin_option = 'YES' THEN ' WITH ADMIN OPTION;' ELSE ';' END
FROM dba_role_privs
WHERE grantee = UPPER('&&SOURCE_USER')
ORDER BY granted_role;

-- SYSTEM PRIVILEGES --
SELECT 'GRANT ' || privilege || ' TO ' || UPPER('&&TARGET_USER') ||
       CASE WHEN admin_option = 'YES' THEN ' WITH ADMIN OPTION;' ELSE ';' END
FROM dba_sys_privs
WHERE grantee = UPPER('&&SOURCE_USER')
ORDER BY privilege;

-- OBJECT PRIVILEGES --
SELECT 'GRANT ' || privilege || ' ON ' || owner || '.' || table_name ||
       ' TO ' || UPPER('&&TARGET_USER') ||
       CASE WHEN grantable = 'YES' THEN ' WITH GRANT OPTION;' ELSE ';' END
FROM dba_tab_privs
WHERE grantee = UPPER('&&SOURCE_USER')
ORDER BY owner, table_name, privilege;

SET FEEDBACK ON HEADING ON VERIFY ON