# Decisions.AD
Custom AD integration project to be used with Decisions (www.decisions.com).
Supports integrated authentication and adds posibility to provide "Additional attributes" that returns any attribute from the AD object.
Integrated authentication only works on Windows. Add ENV variable LDAPTLS_REQCERT: never if you get SSL errors in Linux containers.
Steps that creates users or change the password require LDAPs to work.

Check out the project and build it on your own machine, or download the compiled DLL from Decisions.AD/Zitac.AD.Steps/bin/Debug/net7.0/
Place the DLL file in the Decisions Server\modules\Decisions.Local\CoreServicesDlls folder