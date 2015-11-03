'General Header
'#####################################################################################################################
'Script Description		: Script to Run Individual Business Component
'Test Tool/Version		: HP Quick Test Professional 9.5 and above
'Test Tool Settings		: N.A.
'Application Automated	: N.A.
'Author					: Cognizant
'Date Created			: 07/07/2008
'#####################################################################################################################
Option Explicit	'Forcing Variable declarations

Environment.Value("CurrentScenario") = Parameter("CurrentScenario")
Environment.Value("CurrentTestCase") = Parameter("CurrentTestCase")
Environment.Value("CurrentIteration") = 1
Environment.Value("CurrentSubIteration") = 1

Execute Parameter("CurrentKeyword")