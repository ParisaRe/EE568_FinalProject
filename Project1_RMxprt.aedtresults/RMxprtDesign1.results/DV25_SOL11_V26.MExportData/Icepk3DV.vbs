' ----------------------------------------------------------------------
' Script Created by RMxprt to generate Maxwell 3D Version 2016.0 design 
' Can specify one arg to setup external circuit                         
' ----------------------------------------------------------------------

On Error Resume Next
Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
On Error Goto 0
Set oDesktop = oAnsoftApp.GetAppDesktop()
Set oArgs = AnsoftScript.arguments
oDesktop.RestoreWindow
Set oProject = oDesktop.GetActiveProject()
if (oArgs.Count > 0) then 
  Set oDesign = oProject.GetDesign(oArgs(0))
else
  Set oDesign = oProject.GetActiveDesign()
end if
designName = oDesign.GetName
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", "mm", _
  "Rescale:=", false)
oDesign.SetSolutionType "Transient"
Set oModule = oDesign.GetModule("BoundarySetup")
oEditor.SetModelValidationSettings Array("NAME:Validation Options", _
  "EntityCheckLevel:=", "Strict", "IgnoreUnclassifiedObjects:=", true)
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
  "InsulatorThreshold:=", 2.5e+06)
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:fractions", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "1"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:halfAxial", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "0"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:endRegion", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "296.4005338297477mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:Di", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "5484.975692mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:Do", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "6094.417435mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:L", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "1076.972459mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:slot_h", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "102.077952mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:slot_w", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "22.63265393mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:d_o", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "5454.808326mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:d_i", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "4000mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:pole_w", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "360.2540796mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:pole_h", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "358.4315657mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:shoe_w", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "450.3175995mm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:shoe_h", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "98.78203737mm"))))
On Error Goto 0
Set oDefinitionManager = oProject.GetDefinitionManager()
oDefinitionManager.ModifyLibraries designName, Array("NAME:PersonalLib"), _
  Array("NAME:UserLib"), Array("NAME:SystemLib", "Materials:=", Array(_
  "Materials", "RMxprt"))
if (oDefinitionManager.DoesMaterialExist("Cogent_3DSF0.950")) then
else
oDefinitionManager.AddMaterial Array("NAME:Cogent_3DSF0.950", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 5, "Y:=", 0.00848205), Array("NAME:Coordinate", _
  "X:=", 10, "Y:=", 0.0186535), Array("NAME:Coordinate", "X:=", 20, "Y:=", _
  0.0452376), Array("NAME:Coordinate", "X:=", 31.5, "Y:=", 0.1), Array(_
  "NAME:Coordinate", "X:=", 42, "Y:=", 0.2), Array("NAME:Coordinate", "X:=", _
  49.4, "Y:=", 0.3), Array("NAME:Coordinate", "X:=", 56.1, "Y:=", 0.4), _
  Array("NAME:Coordinate", "X:=", 63.1, "Y:=", 0.5), Array("NAME:Coordinate", _
  "X:=", 70.7, "Y:=", 0.6), Array("NAME:Coordinate", "X:=", 79.5, "Y:=", 0.7), _
  Array("NAME:Coordinate", "X:=", 90.1, "Y:=", 0.8), Array("NAME:Coordinate", _
  "X:=", 103, "Y:=", 0.9), Array("NAME:Coordinate", "X:=", 121, "Y:=", 1), _
  Array("NAME:Coordinate", "X:=", 145, "Y:=", 1.1), Array("NAME:Coordinate", _
  "X:=", 185, "Y:=", 1.2), Array("NAME:Coordinate", "X:=", 273, "Y:=", 1.3), _
  Array("NAME:Coordinate", "X:=", 557, "Y:=", 1.4), Array("NAME:Coordinate", _
  "X:=", 1520, "Y:=", 1.5), Array("NAME:Coordinate", "X:=", 3560, "Y:=", 1.6), _
  Array("NAME:Coordinate", "X:=", 6730, "Y:=", 1.7), Array("NAME:Coordinate", _
  "X:=", 11400, "Y:=", 1.8), Array("NAME:Coordinate", "X:=", 20000, "Y:=", _
  1.89285), Array("NAME:Coordinate", "X:=", 40000, "Y:=", 2.0116), Array(_
  "NAME:Coordinate", "X:=", 75000, "Y:=", 2.11928), Array("NAME:Coordinate", _
  "X:=", 130000, "Y:=", 2.21351), Array("NAME:Coordinate", "X:=", 236790, _
  "Y:=", 2.3745), Array("NAME:Coordinate", "X:=", 431300, "Y:=", 2.6336), _
  Array("NAME:Coordinate", "X:=", 785590, "Y:=", 3.0869), Array(_
  "NAME:Coordinate", "X:=", 1.4309e+06, "Y:=", 3.9022), Array(_
  "NAME:Coordinate", "X:=", 7.884e+06, "Y:=", 12.0552))), Array(_
  "NAME:core_loss_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Electrical Steel"), "core_loss_kh:=", 216.205, "core_loss_kc:=", 0.595593, _
  "core_loss_ke:=", 2.66671, "conductivity:=", 1.81818e+06, "mass_density:=", _
  7650, Array("NAME:stacking_type", "property_type:=", "ChoiceProperty", _
  "Choice:=", "Lamination"), "stacking_factor:=", "0.95", Array(_
  "NAME:stacking_direction", "property_type:=", "ChoiceProperty", "Choice:=", _
  "V(3)"))
end if
if (oDefinitionManager.DoesMaterialExist("Cogent_3DSF0.970")) then
else
oDefinitionManager.AddMaterial Array("NAME:Cogent_3DSF0.970", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 159.2, "Y:=", 0.2402), Array("NAME:Coordinate", _
  "X:=", 318.3, "Y:=", 0.8654), Array("NAME:Coordinate", "X:=", 477.5, "Y:=", _
  1.1106), Array("NAME:Coordinate", "X:=", 636.6, "Y:=", 1.2458), Array(_
  "NAME:Coordinate", "X:=", 795.8, "Y:=", 1.331), Array("NAME:Coordinate", _
  "X:=", 1591.5, "Y:=", 1.5), Array("NAME:Coordinate", "X:=", 3183.1, "Y:=", _
  1.6), Array("NAME:Coordinate", "X:=", 4774.6, "Y:=", 1.683), Array(_
  "NAME:Coordinate", "X:=", 6366.2, "Y:=", 1.741), Array("NAME:Coordinate", _
  "X:=", 7957.7, "Y:=", 1.78), Array("NAME:Coordinate", "X:=", 15915.5, "Y:=", _
  1.905), Array("NAME:Coordinate", "X:=", 31831, "Y:=", 2.025), Array(_
  "NAME:Coordinate", "X:=", 47746.5, "Y:=", 2.085), Array("NAME:Coordinate", _
  "X:=", 63662, "Y:=", 2.13), Array("NAME:Coordinate", "X:=", 79577.5, "Y:=", _
  2.165), Array("NAME:Coordinate", "X:=", 159155, "Y:=", 2.28), Array(_
  "NAME:Coordinate", "X:=", 318310, "Y:=", 2.485), Array("NAME:Coordinate", _
  "X:=", 397887, "Y:=", 2.5851), Array("NAME:Coordinate", "X:=", 1.19366e+06, _
  "Y:=", 3.5861))), "conductivity:=", 2e+06, "mass_density:=", 7872, Array(_
  "NAME:stacking_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Lamination"), "stacking_factor:=", "0.97", Array("NAME:stacking_direction", _
  "property_type:=", "ChoiceProperty", "Choice:=", "V(3)"))
end if
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5469.892009mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", "Name:=", _
  "Length", "Value:=", "1669.773526659495mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", _
  "Value:=", "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "Shaft", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "DiaGap", "(Di + d_o)/2"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "Length", "max(L, L) + 2*endRegion"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5469.892009mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "6094.417435mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1669.773526659495mm"), Array("NAME:Pair", _
  "Name:=", "SegAngle", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", _
  "Fractions", "Value:=", "1"), Array("NAME:Pair", "Name:=", "HalfAxial", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), _
  Array("NAME:Attributes", "Name:=", "OuterRegion", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "DiaGap", "(Di + d_o)/2"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Length", "max(L, L) + 2*endRegion"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "DiaYoke", "Do"
On Error Goto 0
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Fractions", "fractions"
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "HalfAxial", "halfAxial"
On Error Goto 0
oEditor.Copy Array("NAME:Selections", "Selections:=", "OuterRegion")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "1"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "Tool"))))
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.SetSymmetryMultiplier "fractions*(1+halfAxial)"
Set oModule = oDesign.GetModule("BoundarySetup")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SlotCore", "Version:=", "12.1", "NoOfParameters:=", 19, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5484.975692mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "6094.417435mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1076.972459mm"), Array("NAME:Pair", _
  "Name:=", "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", _
  "Value:=", "252"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), _
  Array("NAME:Pair", "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", _
  "102.077952mm"), Array("NAME:Pair", "Name:=", "Bs0", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "HalfSlot", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", _
  "LenRegion", "Value:=", "1669.773526659495mm"), Array("NAME:Pair", "Name:=", _
  "InfoCore", "Value:=", "0"))), Array("NAME:Attributes", "Name:=", "Stator", _
  "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "Cogent_3DSF0.950", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "DiaGap", "Di"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "DiaYoke", "Do"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "Length", "L"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "Hs2", "slot_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "Bs1", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "Bs2", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "LenRegion", "max(L, L) + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SlotCore", "Version:=", "12.1", "NoOfParameters:=", 19, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5484.975692mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "6094.417435mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1076.972459mm"), Array("NAME:Pair", _
  "Name:=", "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", _
  "Value:=", "252"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), _
  Array("NAME:Pair", "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", _
  "102.077952mm"), Array("NAME:Pair", "Name:=", "Bs0", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "HalfSlot", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", _
  "LenRegion", "Value:=", "1669.773526659495mm"), Array("NAME:Pair", "Name:=", _
  "InfoCore", "Value:=", "1"))), Array("NAME:Attributes", "Name:=", _
  "StatorSlotInsu", "Flags:=", "", "Color:=", "(132 132 193)", _
  "Transparency:=", 0, "PartCoordinateSystem:=", "Global", "MaterialName:=", _
  "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "DiaGap", "Di"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "DiaYoke", "Do"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "Length", "L"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "Hs2", "slot_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "Bs1", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "Bs2", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "StatorSlotInsu:CreateUserDefinedPart:1", "LenRegion", _
  "max(L, L) + 2*endRegion"
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:delta", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "-45.907126185930892deg"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:conds", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "1"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:R1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.012288462899128726ohm"))))
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:Le1", "PropType:=", "VariableProp", "UserDef:=", true, "Value:=", _
  "0.00016735776220757695H"))))
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5484.975692mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "6094.417435mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1076.972459mm"), Array("NAME:Pair", _
  "Name:=", "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", _
  "Value:=", "252"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), _
  Array("NAME:Pair", "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", _
  "Name:=", "Hs1", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs2", _
  "Value:=", "102.077952mm"), Array("NAME:Pair", "Name:=", "Bs0", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", _
  "Name:=", "CoilPitch", "Value:=", "8"), Array("NAME:Pair", "Name:=", _
  "EndExt", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SpanExt", _
  "Value:=", "0.1mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "10deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", _
  "1669.773526659495mm"), Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", _
  "1"))), Array("NAME:Attributes", "Name:=", "Coil", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "DiaGap", "Di"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "DiaYoke", "Do"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "Length", "L"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "Hs2", "slot_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "Bs1", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "Bs2", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "LenRegion", "max(L, L) + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Coil"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", _
  "1.4285714285714286deg", "NumClones:=", "252"), Array("NAME:Options", _
  "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Coil"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Coil_0"))))
On Error Resume Next
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "StatorSlotInsu", _
  "Tool Parts:=", "Coil_0,Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7" & _
  ",Coil_8,Coil_9,Coil_10,Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16" & _
  ",Coil_17,Coil_18,Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24" & _
  ",Coil_25,Coil_26,Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32" & _
  ",Coil_33,Coil_34,Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40" & _
  ",Coil_41,Coil_42,Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48" & _
  ",Coil_49,Coil_50,Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56" & _
  ",Coil_57,Coil_58,Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64" & _
  ",Coil_65,Coil_66,Coil_67,Coil_68,Coil_69,Coil_70,Coil_71,Coil_72" & _
  ",Coil_73,Coil_74,Coil_75,Coil_76,Coil_77,Coil_78,Coil_79,Coil_80" & _
  ",Coil_81,Coil_82,Coil_83,Coil_84,Coil_85,Coil_86,Coil_87,Coil_88" & _
  ",Coil_89,Coil_90,Coil_91,Coil_92,Coil_93,Coil_94,Coil_95,Coil_96" & _
  ",Coil_97,Coil_98,Coil_99,Coil_100,Coil_101,Coil_102,Coil_103,Coil_104" & _
  ",Coil_105,Coil_106,Coil_107,Coil_108,Coil_109,Coil_110,Coil_111" & _
  ",Coil_112,Coil_113,Coil_114,Coil_115,Coil_116,Coil_117,Coil_118" & _
  ",Coil_119,Coil_120,Coil_121,Coil_122,Coil_123,Coil_124,Coil_125" & _
  ",Coil_126,Coil_127,Coil_128,Coil_129,Coil_130,Coil_131,Coil_132" & _
  ",Coil_133,Coil_134,Coil_135,Coil_136,Coil_137,Coil_138,Coil_139" & _
  ",Coil_140,Coil_141,Coil_142,Coil_143,Coil_144,Coil_145,Coil_146" & _
  ",Coil_147,Coil_148,Coil_149,Coil_150,Coil_151,Coil_152,Coil_153" & _
  ",Coil_154,Coil_155,Coil_156,Coil_157,Coil_158,Coil_159,Coil_160" & _
  ",Coil_161,Coil_162,Coil_163,Coil_164,Coil_165,Coil_166,Coil_167" & _
  ",Coil_168,Coil_169,Coil_170,Coil_171,Coil_172,Coil_173,Coil_174" & _
  ",Coil_175,Coil_176,Coil_177,Coil_178,Coil_179,Coil_180,Coil_181" & _
  ",Coil_182,Coil_183,Coil_184,Coil_185,Coil_186,Coil_187,Coil_188" & _
  ",Coil_189,Coil_190,Coil_191,Coil_192,Coil_193,Coil_194,Coil_195" & _
  ",Coil_196,Coil_197,Coil_198,Coil_199,Coil_200,Coil_201,Coil_202" & _
  ",Coil_203,Coil_204,Coil_205,Coil_206,Coil_207,Coil_208,Coil_209" & _
  ",Coil_210,Coil_211,Coil_212,Coil_213,Coil_214,Coil_215,Coil_216" & _
  ",Coil_217,Coil_218,Coil_219,Coil_220,Coil_221,Coil_222,Coil_223" & _
  ",Coil_224,Coil_225,Coil_226,Coil_227,Coil_228,Coil_229,Coil_230" & _
  ",Coil_231,Coil_232,Coil_233,Coil_234,Coil_235,Coil_236,Coil_237" & _
  ",Coil_238,Coil_239,Coil_240,Coil_241,Coil_242,Coil_243,Coil_244" & _
  ",Coil_245,Coil_246,Coil_247,Coil_248,Coil_249,Coil_250,Coil_251"), _
  Array("NAME:SubtractParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", true)
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "StatorSlotInsu", _
  "Tool Parts:=", "Stator"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", true)
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5454.808326mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1076.972459mm"), Array("NAME:Pair", _
  "Name:=", "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "Hs0", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "0deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "450.3175995mm"), Array("NAME:Pair", "Name:=", "HeightShoe", _
  "Value:=", "98.78203737mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "360.2540796mm"), Array("NAME:Pair", "Name:=", "HeightBody", _
  "Value:=", "358.4315657mm"), Array("NAME:Pair", "Name:=", "AirGap2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "CoilEndExt", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "EndRingType", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "RingLength", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingHeight", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", _
  "LenRegion", "Value:=", "1669.773526659495mm"), Array("NAME:Pair", "Name:=", _
  "InfoCore", "Value:=", "7"))), Array("NAME:Attributes", "Name:=", "Rotor", _
  "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "Cogent_3DSF0.950", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "DiaGap", "d_o"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "Length", "L"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "WidthShoe", "shoe_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "HeightShoe", "shoe_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "WidthBody", "pole_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "HeightBody", "pole_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "LenRegion", "max(L, L) + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5454.808326mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1076.972459mm"), Array("NAME:Pair", _
  "Name:=", "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "Hs0", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "0deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "450.3175995mm"), Array("NAME:Pair", "Name:=", "HeightShoe", _
  "Value:=", "98.78203737mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "360.2540796mm"), Array("NAME:Pair", "Name:=", "HeightBody", _
  "Value:=", "358.4315657mm"), Array("NAME:Pair", "Name:=", "AirGap2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "CoilEndExt", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "EndRingType", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "RingLength", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingHeight", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", _
  "LenRegion", "Value:=", "1669.773526659495mm"), Array("NAME:Pair", "Name:=", _
  "InfoCore", "Value:=", "6"))), Array("NAME:Attributes", "Name:=", "Pole", _
  "Flags:=", "", "Color:=", "(132 132 193)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "Cogent_3DSF0.970", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "DiaGap", "d_o"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "Length", "L"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "WidthShoe", "shoe_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "HeightShoe", "shoe_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "WidthBody", "pole_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "HeightBody", "pole_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "LenRegion", "max(L, L) + 2*endRegion"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5454.808326mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "1076.972459mm"), Array("NAME:Pair", _
  "Name:=", "Skew", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "Hs0", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Hs01", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Bs0", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs1", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "SlotPitch", "Value:=", "0deg"), Array("NAME:Pair", _
  "Name:=", "CenterPitch", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", _
  "Poles", "Value:=", "32"), Array("NAME:Pair", "Name:=", "WidthShoe", _
  "Value:=", "450.3175995mm"), Array("NAME:Pair", "Name:=", "HeightShoe", _
  "Value:=", "98.78203737mm"), Array("NAME:Pair", "Name:=", "WidthBody", _
  "Value:=", "360.2540796mm"), Array("NAME:Pair", "Name:=", "HeightBody", _
  "Value:=", "358.4315657mm"), Array("NAME:Pair", "Name:=", "AirGap2", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Offset", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "CoilEndExt", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "EndRingType", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "RingLength", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingHeight", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "SegAngle", "Value:=", "30deg"), Array("NAME:Pair", "Name:=", _
  "LenRegion", "Value:=", "1669.773526659495mm"), Array("NAME:Pair", "Name:=", _
  "InfoCore", "Value:=", "2"))), Array("NAME:Attributes", "Name:=", "Field", _
  "Flags:=", "", "Color:=", "(255 128 192)", "Transparency:=", 0, _
  "PartCoordinateSystem:=", "Global", "MaterialName:=", "copper", _
  "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "DiaGap", "d_o"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "Length", "L"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "WidthShoe", "shoe_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "HeightShoe", "shoe_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "WidthBody", "pole_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "HeightBody", "pole_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "LenRegion", "max(L, L) + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "Field"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "11.25deg", _
  "NumClones:=", "32"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Field"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "Field_0"))))
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion"), _
  Array("NAME:ChangedProps", Array("NAME:Transparent", "Value:=", 0.75))))
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Shaft,Stator,Coil_0" & _
  ",Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7,Coil_8,Coil_9,Coil_10" & _
  ",Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16,Coil_17,Coil_18" & _
  ",Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24,Coil_25,Coil_26" & _
  ",Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32,Coil_33,Coil_34" & _
  ",Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40,Coil_41,Coil_42" & _
  ",Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48,Coil_49,Coil_50" & _
  ",Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56,Coil_57,Coil_58" & _
  ",Coil_59,Coil_60,Coil_61,Coil_62,Coil_63,Coil_64,Coil_65,Coil_66" & _
  ",Coil_67,Coil_68,Coil_69,Coil_70,Coil_71,Coil_72,Coil_73,Coil_74" & _
  ",Coil_75,Coil_76,Coil_77,Coil_78,Coil_79,Coil_80,Coil_81,Coil_82" & _
  ",Coil_83,Coil_84,Coil_85,Coil_86,Coil_87,Coil_88,Coil_89,Coil_90" & _
  ",Coil_91,Coil_92,Coil_93,Coil_94,Coil_95,Coil_96,Coil_97,Coil_98" & _
  ",Coil_99,Coil_100,Coil_101,Coil_102,Coil_103,Coil_104,Coil_105,Coil_106" & _
  ",Coil_107,Coil_108,Coil_109,Coil_110,Coil_111,Coil_112,Coil_113" & _
  ",Coil_114,Coil_115,Coil_116,Coil_117,Coil_118,Coil_119,Coil_120" & _
  ",Coil_121,Coil_122,Coil_123,Coil_124,Coil_125,Coil_126,Coil_127" & _
  ",Coil_128,Coil_129,Coil_130,Coil_131,Coil_132,Coil_133,Coil_134" & _
  ",Coil_135,Coil_136,Coil_137,Coil_138,Coil_139,Coil_140,Coil_141" & _
  ",Coil_142,Coil_143,Coil_144,Coil_145,Coil_146,Coil_147,Coil_148" & _
  ",Coil_149,Coil_150,Coil_151,Coil_152,Coil_153,Coil_154,Coil_155" & _
  ",Coil_156,Coil_157,Coil_158,Coil_159,Coil_160,Coil_161,Coil_162" & _
  ",Coil_163,Coil_164,Coil_165,Coil_166,Coil_167,Coil_168,Coil_169" & _
  ",Coil_170,Coil_171,Coil_172,Coil_173,Coil_174,Coil_175,Coil_176" & _
  ",Coil_177,Coil_178,Coil_179,Coil_180,Coil_181,Coil_182,Coil_183" & _
  ",Coil_184,Coil_185,Coil_186,Coil_187,Coil_188,Coil_189,Coil_190" & _
  ",Coil_191,Coil_192,Coil_193,Coil_194,Coil_195,Coil_196,Coil_197" & _
  ",Coil_198,Coil_199,Coil_200,Coil_201,Coil_202,Coil_203,Coil_204" & _
  ",Coil_205,Coil_206,Coil_207,Coil_208,Coil_209,Coil_210,Coil_211" & _
  ",Coil_212,Coil_213,Coil_214,Coil_215,Coil_216,Coil_217,Coil_218" & _
  ",Coil_219,Coil_220,Coil_221,Coil_222,Coil_223,Coil_224,Coil_225" & _
  ",Coil_226,Coil_227,Coil_228,Coil_229,Coil_230,Coil_231,Coil_232" & _
  ",Coil_233,Coil_234,Coil_235,Coil_236,Coil_237,Coil_238,Coil_239" & _
  ",Coil_240,Coil_241,Coil_242,Coil_243,Coil_244,Coil_245,Coil_246" & _
  ",Coil_247,Coil_248,Coil_249,Coil_250,Coil_251,Rotor,Pole,Field_0" & _
  ",Field_1,Field_2,Field_3,Field_4,Field_5,Field_6,Field_7,Field_8" & _
  ",Field_9,Field_10,Field_11,Field_12,Field_13,Field_14,Field_15,Field_16" & _
  ",Field_17,Field_18,Field_19,Field_20,Field_21,Field_22,Field_23" & _
  ",Field_24,Field_25,Field_26,Field_27,Field_28,Field_29,Field_30" & _
  ",Field_31", "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.FitAll
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "Transient", Array("NAME:Setup1", "StopTime:=", "0.2s", _
  "TimeStep:=", "0.0002s")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetMinimumTimeStep "0.002ms" 
set oUnclassified = oEditor.GetObjectsInGroup("Unclassified")
Dim oObject
For Each oObject in oUnclassified
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", oObject), Array(_
  "NAME:ChangedProps", Array("NAME:Model", "Value:=", false))))
Next
