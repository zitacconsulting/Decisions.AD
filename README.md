# Decisions.AD
Custom AD integration project to be used with Decisions (www.decisions.com).
Supports integrated authentication and adds posibility to provide "Additional attributes" that returns any attribute from the AD object.
Integrated authentication only works on Windows. Add ENV variable LDAPTLS_REQCERT: never if you get SSL errors in Linux containers.
Steps that creates users or change the password require LDAPs to work.