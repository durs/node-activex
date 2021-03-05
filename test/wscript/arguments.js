WScript.Echo("Same count: " + (WScript.Arguments.length==WScript.Arguments.Count()) );

objArgs = WScript.Arguments
WScript.Echo(WScript.Arguments.Count());
for (i=0; i<objArgs.length; i++)
{
    WScript.Echo(objArgs(i))
}