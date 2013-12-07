CREATE TABLE APPLICATIONS
(
  ID serial NOT NULL,
  NAME VARCHAR(800),
  CONSTRAINT APPPK PRIMARY KEY (ID),
  CONSTRAINT APPNAMEUNIQ UNIQUE (NAME)
);
CREATE TABLE CPU
(
  ID serial NOT NULL,
  CPUNAME VARCHAR(255),
  CONSTRAINT CPUPK PRIMARY KEY (ID),
  CONSTRAINT CPUUNIQUE UNIQUE (CPUNAME)
);
CREATE TABLE EXCEPTIONCLASSES
(
  ID serial NOT NULL,
  EXCEPTIONCLASS VARCHAR(800),
  CONSTRAINT EXCEPTIONSPK PRIMARY KEY (ID),
  CONSTRAINT UNQ_EXCEPTIONCLASSES_CLASS UNIQUE (EXCEPTIONCLASS)
);
CREATE TABLE EXCEPTIONMESSAGES
(
  ID serial NOT NULL,
  EXCEPTIONCLASS INTEGER,
  EXCEPTIONMESSAGE VARCHAR(800),
  CONSTRAINT EXCEPTIONMESSAGESPK PRIMARY KEY (ID),
  CONSTRAINT EXCEPTIONMESSAGEUNIQUE UNIQUE (EXCEPTIONMESSAGE)
);
CREATE TABLE METHODNAMES
(
  ID serial NOT NULL,
  NAME VARCHAR(800),
  CONSTRAINT METHODNAMESPK PRIMARY KEY (ID),
  CONSTRAINT METHODNAMESUNIQUENAME UNIQUE (NAME)
);
CREATE TABLE OPTIONS
(
  OPTIONNAME VARCHAR(255) NOT NULL,
  OPTIONVALUE VARCHAR(255),
  REMARKS VARCHAR(255),
  CONSTRAINT OPTIONSPK PRIMARY KEY (OPTIONNAME)
);
CREATE TABLE OS
(
  ID serial NOT NULL,
  OSNAME VARCHAR(255),
  CONSTRAINT OSPK PRIMARY KEY (ID),
  CONSTRAINT OSUNIQUE UNIQUE (OSNAME)
);
CREATE TABLE RESULTVALUES
(
  ID serial NOT NULL,
  NAME VARCHAR(800),
  CONSTRAINT RESULTVALUESPK PRIMARY KEY (ID),
  CONSTRAINT UNQ_RESULTVALUES_NAME UNIQUE (NAME)
);
CREATE TABLE SOURCELOCATIONS
(
  ID serial NOT NULL,
  SOURCEUNIT INTEGER,
  LINE INTEGER,
  CONSTRAINT SOURCELOCATIONSPK PRIMARY KEY (ID),
  CONSTRAINT SOURCELOCATIONSUNIQUE UNIQUE (SOURCEUNIT,LINE)
);
CREATE TABLE SOURCEUNITS
(
  ID serial NOT NULL,
  NAME VARCHAR(800),
  CONSTRAINT SOURCEUNITS_PK PRIMARY KEY (ID),
  CONSTRAINT SOURCEUNITS_NAME_UNIQUE UNIQUE (NAME)
);
CREATE TABLE TESTS
(
  ID serial NOT NULL,
  TESTSUITE INTEGER,
  NAME VARCHAR(800),
  CONSTRAINT TESTSPK PRIMARY KEY (ID),
  CONSTRAINT UNQ_TESTS_SUITENAME UNIQUE (TESTSUITE,NAME)
);
CREATE TABLE TESTRUNS
(
  ID serial NOT NULL,
  DATETIMERAN TIMESTAMP,
	APPLICATIONID INTEGER,
  CPU INTEGER,
  OS INTEGER,
  REVISIONID VARCHAR(800),
  RUNCOMMENT VARCHAR(800),
  TOTALELAPSEDTIME TIME,
  CONSTRAINT TESTRUNSPK PRIMARY KEY (ID)
);
CREATE TABLE TESTRESULTS
(
  ID serial NOT NULL,
  TESTRUN INTEGER,
  TEST INTEGER,
  RESULTVALUE INTEGER,
  EXCEPTIONMESSAGE INTEGER,
  METHODNAME INTEGER,
  SOURCELOCATION INTEGER,
	RESULTCOMMENT VARCHAR(800),
  ELAPSEDTIME TIME,
  CONSTRAINT TESTRESULTSPK PRIMARY KEY (ID)
);
CREATE TABLE TESTSUITES
(
  ID serial NOT NULL,
	PARENTSUITE INTEGER,
  NAME VARCHAR(800),
	DEPTH INTEGER,
  CONSTRAINT TESTSUITESPK PRIMARY KEY (ID),
	CONSTRAINT UNQ_TESTSUITES_NAMEPAR UNIQUE (PARENTSUITE,NAME)
);
CREATE VIEW TESTSUITESFLAT (TESTSUITEID, TESTSUITENAME, DEPTH)
AS  
with recursive suite_tree as (
  select id as testsuiteid, ''||name as testsuitename, depth from TESTSUITES
  where parentsuite is null
  -- to do: find a better way to cast testsuitename from varchar(800) to character varying without limits
union all
  select chi.id as testsuiteid, par.testsuitename||'/'||chi.name as testsuitename, chi.depth from testsuites chi
  join suite_tree par on chi.parentsuite=par.testsuiteid
)
select testsuiteid,testsuitename,depth from suite_tree;

CREATE VIEW FLAT (TESTRUNID, TESTRESULTID, TESTID, APPLICATION, REVISIONID, RUNCOMMENT, TESTRUNDATE, OS, CPU, TESTSUITE, TESTSUITEDEPTH, TESTNAME, TESTRESULT, EXCEPTIONCLASS, EXCEPTIONMESSAGE, METHOD, SOURCELINE, SOURCEUNIT, ELAPSEDTIME)
AS               
SELECT 
R.ID as TESTRUNID, 
TR.ID as TESTRESULTID,
T.ID as TESTID,
AP.NAME as APPLICATION,
R.REVISIONID, 
R.RUNCOMMENT, 
R.DATETIMERAN as TESTRUNDATE,
OS.OSNAME,
CP.CPUNAME,
S.TESTSUITENAME as TESTSUITE,
S.DEPTH as TESTSUITEDEPTH,
T.NAME as TESTNAME,
RV.NAME as RESULT,
E.EXCEPTIONCLASS,
EM.EXCEPTIONMESSAGE as EXCEPTIONMESSAGE,
M.NAME as METHOD,
SL.LINE as SOURCELINE,
SU.NAME as SOURCEUNIT,
TR.ELAPSEDTIME as ELAPSEDTIME
FROM TESTRUNS R inner join TESTRESULTS TR on R.ID=TR.TESTRUN
inner join TESTS T on TR.TEST=T.ID
inner join TESTSUITESFLAT S on T.TESTSUITE=S.TESTSUITEID
inner join RESULTVALUES RV on TR.RESULTVALUE=RV.ID
left join APPLICATIONS AP on R.APPLICATIONID=AP.ID
left join
EXCEPTIONMESSAGES EM on TR.EXCEPTIONMESSAGE=EM.ID
left join EXCEPTIONCLASSES E on EM.EXCEPTIONCLASS=E.ID
left join METHODNAMES M on TR.METHODNAME=M.ID
left join SOURCELOCATIONS SL on TR.SOURCELOCATION=SL.ID
left join SOURCEUNITS SU on SL.SOURCEUNIT=SU.ID
left join OS on R.OS=OS.ID
left join CPU CP on R.CPU=CP.ID;

CREATE VIEW FLATSORTED (TESTRUNID, TESTRESULTID, TESTID, APPLICATION, REVISIONID, RUNCOMMENT, TESTRUNDATE, OS, CPU, TESTSUITE, TESTSUITEDEPTH, TESTNAME, TESTRESULT, EXCEPTIONCLASS, EXCEPTIONMESSAGE, METHOD, SOURCELINE, SOURCEUNIT)
AS    
select
    f.TESTRUNID, f.TESTRESULTID, f.TESTID, 
    f.APPLICATION, f.REVISIONID,
    f.RUNCOMMENT, f.TESTRUNDATE, f.OS, f.CPU,
    f.TESTSUITE, f.TESTSUITEDEPTH, f.TESTNAME, f.TESTRESULT,
    f.EXCEPTIONCLASS, f.EXCEPTIONMESSAGE,
    f.METHOD, f.SOURCELINE, f.SOURCEUNIT from 
flat f
order by f.TESTRUNDATE desc, f.application, f.revisionid, f.TESTSUITEDEPTH, f.TESTSUITE, f.TESTNAME;

ALTER TABLE EXCEPTIONMESSAGES ADD CONSTRAINT FK_EXCEPTIONCLASSES_CLASS
  FOREIGN KEY (EXCEPTIONCLASS) REFERENCES EXCEPTIONCLASSES (ID) ON UPDATE
CASCADE ON DELETE CASCADE;
ALTER TABLE SOURCELOCATIONS ADD CONSTRAINT SOURCELOCATIONSFK_UNIT
  FOREIGN KEY (SOURCEUNIT) REFERENCES SOURCEUNITS (ID) ON UPDATE CASCADE ON
DELETE CASCADE;
ALTER TABLE TESTS ADD CONSTRAINT TESTSTESTSUITESFK
  FOREIGN KEY (TESTSUITE) REFERENCES TESTSUITES (ID) ON UPDATE CASCADE
ON DELETE CASCADE;
ALTER TABLE TESTRUNS ADD CONSTRAINT FK_TESTRUNSCPU
  FOREIGN KEY (CPU) REFERENCES CPU (ID) ON UPDATE CASCADE ON DELETE CASCADE;
ALTER TABLE TESTRUNS ADD CONSTRAINT FK_TESTRUNSOS
  FOREIGN KEY (OS) REFERENCES OS (ID) ON UPDATE CASCADE ON DELETE CASCADE;
ALTER TABLE TESTRUNS ADD CONSTRAINT FK_TESTRUNS_APPLICATIONS
  FOREIGN KEY (APPLICATIONID) REFERENCES APPLICATIONS (ID) ON UPDATE CASCADE ON DELETE CASCADE;	
ALTER TABLE TESTRESULTS ADD CONSTRAINT FK_TESTRES_EXCEPTION
  FOREIGN KEY (EXCEPTIONMESSAGE) REFERENCES EXCEPTIONMESSAGES (ID) ON UPDATE
CASCADE ON DELETE CASCADE;
ALTER TABLE TESTRESULTS ADD CONSTRAINT FK_TESTRES_METHODNAME
  FOREIGN KEY (METHODNAME) REFERENCES METHODNAMES (ID) ON UPDATE CASCADE ON
DELETE CASCADE;
ALTER TABLE TESTRESULTS ADD CONSTRAINT FK_TESTRES_RESULT
  FOREIGN KEY (RESULTVALUE) REFERENCES RESULTVALUES (ID) ON UPDATE CASCADE
ON DELETE CASCADE;
ALTER TABLE TESTRESULTS ADD CONSTRAINT FK_TESTRES_SOURCELOC
  FOREIGN KEY (SOURCELOCATION) REFERENCES SOURCELOCATIONS (ID) ON UPDATE
CASCADE ON DELETE CASCADE;
ALTER TABLE TESTRESULTS ADD CONSTRAINT FK_TESTRES_TEST
  FOREIGN KEY (TEST) REFERENCES TESTS (ID) ON UPDATE CASCADE ON
DELETE CASCADE;
ALTER TABLE TESTRESULTS ADD CONSTRAINT FK_TESTRES_TESTRUN
  FOREIGN KEY (TESTRUN) REFERENCES TESTRUNS (ID) ON UPDATE CASCADE ON DELETE
CASCADE;
ALTER TABLE TESTSUITES ADD CONSTRAINT FK_TESTSUITES_PARENT
  FOREIGN KEY (PARENTSUITE) REFERENCES TESTSUITES (ID) ON UPDATE CASCADE ON DELETE CASCADE;

COMMENT ON TABLE options
  IS 'Stores schema version and any application-specific options.';
COMMENT ON TABLE sourceunits
  IS 'Pascal units where errrors occurred';
COMMENT ON COLUMN testruns.revisionid IS 'String that uniquely identifies
the revision/version of the code that is tested. Useful when running
regression tests, identifying when an error occurred first etc.';
COMMENT ON COLUMN testruns.runcomment IS 'Comment provided by user/test run
suite on this test run (e.g. used compiler flags)';
COMMENT ON COLUMN tests.name IS 'Identifies both the name (and following
the FK), the test suite. This means that multiple test suites with the same
test name text are allowed.';