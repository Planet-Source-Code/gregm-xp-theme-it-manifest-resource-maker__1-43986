
XP Theme-It Manifest Resource Maker

--
Create a custom manifest for any executable you choose, and force the program to load the new XP style controls based on the current XP Style Theme. Add any created manifest to a custom resource file, so it can be loaded to a VB project and compiled directly into the program's executable. Organize and maintain a list of all manifests created on your system so the visual theme can easily be restored if necessary. Also enables you to create a custom manifest by right-clicking any executable from Windows Explorer. If you like this code, please vote for it at PSC.

NOTES:
You need to ensure that any form module using the style theme, initializes the
ComCtl32.dll by calling the ComCtl InitCommonControls
API.

Project Requirements:
Microsoft Resource Compiler - (RC.EXE)

Stream Data Object Library 1.0 - (stmdata.tlb)
Register type library if needed.