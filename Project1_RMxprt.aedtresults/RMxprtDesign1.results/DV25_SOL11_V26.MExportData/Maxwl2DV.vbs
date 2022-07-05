' ----------------------------------------------------------------------
' Script Created by RMxprt to generate Maxwell 2D Version 2016.0 design 
' Can specify one arg to setup external circuit                         
' ----------------------------------------------------------------------

On Error Resume Next
Set oAnsoftApp = CreateObject("Ansoft.ElectronicsDesktop")
On Error Goto 0
Set oDesktop = oAnsoftApp.GetAppDesktop()
Set oArgs = AnsoftScript.arguments
oDesktop.RestoreWindow
Set oProject = oDesktop.GetActiveProject()
Set oDesign = oProject.GetActiveDesign()
designName = oDesign.GetName
Set oEditor = oDesign.SetActiveEditor("3D Modeler")
oEditor.SetModelUnits Array("NAME:Units Parameter", "Units:=", "mm", _
  "Rescale:=", false)
oDesign.SetSolutionType "Transient", "XY"
Set oModule = oDesign.GetModule("BoundarySetup")
if (oArgs.Count = 1) then 
oModule.EditExternalCircuit oArgs(0), Array(), Array(), Array(), Array()
end if
oEditor.SetModelValidationSettings Array("NAME:Validation Options", _
  "EntityCheckLevel:=", "Strict", "IgnoreUnclassifiedObjects:=", true)
On Error Resume Next
Set oModule = oDesign.GetModule("MeshSetup")
oModule.InitialMeshSettings Array("NAME:MeshSettings", "MeshMethod:=", _
  "AnsoftTAU")
On Error Goto 0
On Error Resume Next
oDesign.ChangeProperty Array("NAME:AllTabs", Array("NAME:LocalVariableTab", _
  Array("NAME:PropServers", "LocalVariables"), Array("NAME:NewProps", Array(_
  "NAME:fractions", "PropType:=", "VariableProp", "UserDef:=", true, _
  "Value:=", "4"))))
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
if (oDefinitionManager.DoesMaterialExist("Cogent_2DSF0.950")) then
else
oDefinitionManager.AddMaterial Array("NAME:Cogent_2DSF0.950", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 5, "Y:=", 0.00805826), Array("NAME:Coordinate", _
  "X:=", 10, "Y:=", 0.0177215), Array("NAME:Coordinate", "X:=", 20, "Y:=", _
  0.042977), Array("NAME:Coordinate", "X:=", 31.5, "Y:=", 0.095002), Array(_
  "NAME:Coordinate", "X:=", 42, "Y:=", 0.190003), Array("NAME:Coordinate", _
  "X:=", 49.4, "Y:=", 0.285003), Array("NAME:Coordinate", "X:=", 56.1, "Y:=", _
  0.380004), Array("NAME:Coordinate", "X:=", 63.1, "Y:=", 0.475004), Array(_
  "NAME:Coordinate", "X:=", 70.7, "Y:=", 0.570004), Array("NAME:Coordinate", _
  "X:=", 79.5, "Y:=", 0.665005), Array("NAME:Coordinate", "X:=", 90.1, "Y:=", _
  0.760006), Array("NAME:Coordinate", "X:=", 103, "Y:=", 0.855006), Array(_
  "NAME:Coordinate", "X:=", 121, "Y:=", 0.950008), Array("NAME:Coordinate", _
  "X:=", 145, "Y:=", 1.04501), Array("NAME:Coordinate", "X:=", 185, "Y:=", _
  1.14001), Array("NAME:Coordinate", "X:=", 273, "Y:=", 1.23502), Array(_
  "NAME:Coordinate", "X:=", 557, "Y:=", 1.33003), Array("NAME:Coordinate", _
  "X:=", 1520, "Y:=", 1.4251), Array("NAME:Coordinate", "X:=", 3560, "Y:=", _
  1.52022), Array("NAME:Coordinate", "X:=", 6730, "Y:=", 1.61542), Array(_
  "NAME:Coordinate", "X:=", 11400, "Y:=", 1.71072), Array("NAME:Coordinate", _
  "X:=", 20000, "Y:=", 1.79946), Array("NAME:Coordinate", "X:=", 40000, "Y:=", _
  1.91353), Array("NAME:Coordinate", "X:=", 75000, "Y:=", 2.01803), Array(_
  "NAME:Coordinate", "X:=", 130000, "Y:=", 2.111), Array("NAME:Coordinate", _
  "X:=", 236790, "Y:=", 2.27065), Array("NAME:Coordinate", "X:=", 431300, _
  "Y:=", 2.52902), Array("NAME:Coordinate", "X:=", 785590, "Y:=", 2.98192), _
  Array("NAME:Coordinate", "X:=", 1.4309e+06, "Y:=", 3.797), Array(_
  "NAME:Coordinate", "X:=", 7.884e+06, "Y:=", 11.9478))), Array(_
  "NAME:core_loss_type", "property_type:=", "ChoiceProperty", "Choice:=", _
  "Electrical Steel"), "core_loss_kh:=", 227.584, "core_loss_kc:=", 0.62694, _
  "core_loss_ke:=", 2.73599, "conductivity:=", 1.81818e+06, "mass_density:=", _
  7267.5) 
end if
if (oDefinitionManager.DoesMaterialExist("Cogent_2DSF0.970")) then
else
oDefinitionManager.AddMaterial Array("NAME:Cogent_2DSF0.970", _
  "CoordinateSystemType:=", "Cartesian", Array("NAME:AttachedData"), Array(_
  "NAME:ModifierData"), Array("NAME:permeability", "property_type:=", _
  "nonlinear", "HUnit:=", "A_per_meter", "BUnit:=", "tesla", Array(_
  "NAME:BHCoordinates", Array("NAME:Coordinate", "X:=", 0, "Y:=", 0), Array(_
  "NAME:Coordinate", "X:=", 159.2, "Y:=", 0.233), Array("NAME:Coordinate", _
  "X:=", 318.3, "Y:=", 0.83945), Array("NAME:Coordinate", "X:=", 477.5, "Y:=", _
  1.0773), Array("NAME:Coordinate", "X:=", 636.6, "Y:=", 1.20845), Array(_
  "NAME:Coordinate", "X:=", 795.8, "Y:=", 1.2911), Array("NAME:Coordinate", _
  "X:=", 1591.5, "Y:=", 1.45506), Array("NAME:Coordinate", "X:=", 3183.1, _
  "Y:=", 1.55212), Array("NAME:Coordinate", "X:=", 4774.6, "Y:=", 1.63269), _
  Array("NAME:Coordinate", "X:=", 6366.2, "Y:=", 1.68901), Array(_
  "NAME:Coordinate", "X:=", 7957.7, "Y:=", 1.7269), Array("NAME:Coordinate", _
  "X:=", 15915.5, "Y:=", 1.84845), Array("NAME:Coordinate", "X:=", 31831, _
  "Y:=", 1.96545), Array("NAME:Coordinate", "X:=", 47746.5, "Y:=", 2.02425), _
  Array("NAME:Coordinate", "X:=", 63662, "Y:=", 2.0685), Array(_
  "NAME:Coordinate", "X:=", 79577.5, "Y:=", 2.10305), Array("NAME:Coordinate", _
  "X:=", 159155, "Y:=", 2.2176), Array("NAME:Coordinate", "X:=", 318310, _
  "Y:=", 2.42245), Array("NAME:Coordinate", "X:=", 397887, "Y:=", 2.52255), _
  Array("NAME:Coordinate", "X:=", 1.19366e+06, "Y:=", 3.52352))), _
  "conductivity:=", 2e+06, "mass_density:=", 7635.84) 
end if
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5469.892009mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", "Name:=", _
  "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), Array(_
  "NAME:Attributes", "Name:=", "Band", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Band:CreateUserDefinedPart:1", _
  "DiaGap", "(Di + d_o)/2"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Band:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5469.892009mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", "Name:=", _
  "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", "Value:=", _
  "1"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "Shaft", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "DiaGap", "(Di + d_o)/2"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Shaft:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
On Error Goto 0
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/Band", "Version:=", "12.1", "NoOfParameters:=", 7, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5469.892009mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "6094.417435mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Fractions", _
  "Value:=", "4"), Array("NAME:Pair", "Name:=", "HalfAxial", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "100"))), Array(_
  "NAME:Attributes", "Name:=", "OuterRegion", "Flags:=", "", "Color:=", _
  "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "DiaGap", "(Di + d_o)/2"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "DiaYoke", "Do"
On Error Goto 0
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion:CreateUserDefinedPart:1", "Fractions", "fractions"
oEditor.Copy Array("NAME:Selections", "Selections:=", "OuterRegion")
oEditor.Paste
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "OuterRegion1:CreateUserDefinedPart:1", "InfoCore", "1"
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "OuterRegion1"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "Tool"))))
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.SetSymmetryMultiplier "fractions"
Set oModule = oDesign.GetModule("BoundarySetup")
edgeID = oEditor.GetEdgeByPosition(Array("NAME:Parameters", "BodyName:=", _
  "OuterRegion", "XPosition:=", "2154.701947835013mm", "YPosition:=", _
  "2154.701947835013mm", "ZPosition:=", "0mm"))
oModule.AssignVectorPotential Array("NAME:VectorPotential1", "Edges:=", Array(edgeID), _
  "Value:=", "0", "CoordinateSystem:=", "")
edgeID = oEditor.GetEdgeByPosition(Array("NAME:Parameters", "BodyName:=", _
  "OuterRegion", "XPosition:=", "1523.6043587500001mm", "YPosition:=", "0mm", _
  "ZPosition:=", "0mm"))
oModule.AssignMaster Array("NAME:Independent1", "Edges:=", Array(edgeID), _
  "ReverseV:=", false)
edgeID = oEditor.GetEdgeByPosition(Array("NAME:Parameters", "BodyName:=", _
  "OuterRegion", "XPosition:=", "9.3293860055507155e-14mm", "YPosition:=", _
  "1523.6043587500001mm", "ZPosition:=", "0mm"))
oModule.AssignSlave Array("NAME:Dependent1", "Edges:=", Array(edgeID), _
  "ReverseU:=", true, "Master:=", "Independent1", "SameAsMaster:=", true)
oDesign.SetDesignSettings Array("NAME:Design Settings Data", "ModelDepth:=", _
  "1076.97mm")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SlotCore", "Version:=", "12.1", "NoOfParameters:=", 19, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5484.975692mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "6094.417435mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "252"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "1mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "102.077952mm"), Array(_
  "NAME:Pair", "Name:=", "Bs0", "Value:=", "22.63265393mm"), Array(_
  "NAME:Pair", "Name:=", "Bs1", "Value:=", "22.63265393mm"), Array(_
  "NAME:Pair", "Name:=", "Bs2", "Value:=", "22.63265393mm"), Array(_
  "NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", "HalfSlot", _
  "Value:=", "0"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "0"))), Array(_
  "NAME:Attributes", "Name:=", "Stator", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "Cogent_2DSF0.950", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "DiaGap", "Di"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "DiaYoke", "Do"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "Hs2", "slot_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "Bs1", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "Bs2", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Stator:CreateUserDefinedPart:1", _
  "LenRegion", "max(L, L) + 2*endRegion"
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
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "252"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", _
  "102.077952mm"), Array("NAME:Pair", "Name:=", "Bs0", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", _
  "Name:=", "CoilPitch", "Value:=", "8"), Array("NAME:Pair", "Name:=", _
  "EndExt", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SpanExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "10deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", "2"))), Array(_
  "NAME:Attributes", "Name:=", "Coil", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "DiaGap", "Di"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Coil:CreateUserDefinedPart:1", _
  "DiaYoke", "Do"
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
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_0,Coil_63,Coil_126" & _
  ",Coil_189"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_1,Coil_64,Coil_127" & _
  ",Coil_190"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_2,Coil_65,Coil_128" & _
  ",Coil_191"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_3,Coil_66,Coil_129" & _
  ",Coil_192"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_4,Coil_67,Coil_130" & _
  ",Coil_193"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_5,Coil_68,Coil_131" & _
  ",Coil_194"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_6,Coil_69,Coil_132" & _
  ",Coil_195"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_7,Coil_70,Coil_133" & _
  ",Coil_196"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_8,Coil_71,Coil_134" & _
  ",Coil_197"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_9,Coil_72,Coil_135" & _
  ",Coil_198"), Array("NAME:UniteParameters", "CoordinateSystemID:=", -1, _
  "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_10,Coil_73" & _
  ",Coil_136,Coil_199"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_11,Coil_74" & _
  ",Coil_137,Coil_200"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_12,Coil_75" & _
  ",Coil_138,Coil_201"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_13,Coil_76" & _
  ",Coil_139,Coil_202"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_14,Coil_77" & _
  ",Coil_140,Coil_203"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_15,Coil_78" & _
  ",Coil_141,Coil_204"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_16,Coil_79" & _
  ",Coil_142,Coil_205"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_17,Coil_80" & _
  ",Coil_143,Coil_206"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_18,Coil_81" & _
  ",Coil_144,Coil_207"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_19,Coil_82" & _
  ",Coil_145,Coil_208"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_20,Coil_83" & _
  ",Coil_146,Coil_209"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_21,Coil_84" & _
  ",Coil_147,Coil_210"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_22,Coil_85" & _
  ",Coil_148,Coil_211"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_23,Coil_86" & _
  ",Coil_149,Coil_212"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_24,Coil_87" & _
  ",Coil_150,Coil_213"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_25,Coil_88" & _
  ",Coil_151,Coil_214"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_26,Coil_89" & _
  ",Coil_152,Coil_215"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_27,Coil_90" & _
  ",Coil_153,Coil_216"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_28,Coil_91" & _
  ",Coil_154,Coil_217"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_29,Coil_92" & _
  ",Coil_155,Coil_218"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_30,Coil_93" & _
  ",Coil_156,Coil_219"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_31,Coil_94" & _
  ",Coil_157,Coil_220"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_32,Coil_95" & _
  ",Coil_158,Coil_221"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_33,Coil_96" & _
  ",Coil_159,Coil_222"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_34,Coil_97" & _
  ",Coil_160,Coil_223"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_35,Coil_98" & _
  ",Coil_161,Coil_224"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_36,Coil_99" & _
  ",Coil_162,Coil_225"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_37,Coil_100" & _
  ",Coil_163,Coil_226"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_38,Coil_101" & _
  ",Coil_164,Coil_227"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_39,Coil_102" & _
  ",Coil_165,Coil_228"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_40,Coil_103" & _
  ",Coil_166,Coil_229"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_41,Coil_104" & _
  ",Coil_167,Coil_230"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_42,Coil_105" & _
  ",Coil_168,Coil_231"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_43,Coil_106" & _
  ",Coil_169,Coil_232"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_44,Coil_107" & _
  ",Coil_170,Coil_233"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_45,Coil_108" & _
  ",Coil_171,Coil_234"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_46,Coil_109" & _
  ",Coil_172,Coil_235"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_47,Coil_110" & _
  ",Coil_173,Coil_236"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_48,Coil_111" & _
  ",Coil_174,Coil_237"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_49,Coil_112" & _
  ",Coil_175,Coil_238"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_50,Coil_113" & _
  ",Coil_176,Coil_239"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_51,Coil_114" & _
  ",Coil_177,Coil_240"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_52,Coil_115" & _
  ",Coil_178,Coil_241"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_53,Coil_116" & _
  ",Coil_179,Coil_242"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_54,Coil_117" & _
  ",Coil_180,Coil_243"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_55,Coil_118" & _
  ",Coil_181,Coil_244"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_56,Coil_119" & _
  ",Coil_182,Coil_245"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_57,Coil_120" & _
  ",Coil_183,Coil_246"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_58,Coil_121" & _
  ",Coil_184,Coil_247"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_59,Coil_122" & _
  ",Coil_185,Coil_248"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_60,Coil_123" & _
  ",Coil_186,Coil_249"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_61,Coil_124" & _
  ",Coil_187,Coil_250"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Coil_62,Coil_125" & _
  ",Coil_188,Coil_251"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/LapCoil", "Version:=", "12.0", "NoOfParameters:=", 21, _
  "Library:=", "syslib", Array("NAME:ParamVector", Array("NAME:Pair", _
  "Name:=", "DiaGap", "Value:=", "5484.975692mm"), Array("NAME:Pair", _
  "Name:=", "DiaYoke", "Value:=", "6094.417435mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "252"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "6"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs1", _
  "Value:=", "1mm"), Array("NAME:Pair", "Name:=", "Hs2", "Value:=", _
  "102.077952mm"), Array("NAME:Pair", "Name:=", "Bs0", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs1", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", _
  "22.63265393mm"), Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "FilletType", "Value:=", "0"), Array(_
  "NAME:Pair", "Name:=", "Layers", "Value:=", "2"), Array("NAME:Pair", _
  "Name:=", "CoilPitch", "Value:=", "8"), Array("NAME:Pair", "Name:=", _
  "EndExt", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SpanExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "SegAngle", "Value:=", _
  "10deg"), Array("NAME:Pair", "Name:=", "LenRegion", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "InfoCoil", "Value:=", "3"))), Array(_
  "NAME:Attributes", "Name:=", "CoilRe", "Flags:=", "", "Color:=", _
  "(250 167 14)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "CoilRe:CreateUserDefinedPart:1", _
  "DiaGap", "Di"
oEditor.SetPropertyValue "Geometry3DCmdTab", "CoilRe:CreateUserDefinedPart:1", _
  "DiaYoke", "Do"
oEditor.SetPropertyValue "Geometry3DCmdTab", "CoilRe:CreateUserDefinedPart:1", _
  "Hs2", "slot_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "CoilRe:CreateUserDefinedPart:1", _
  "Bs1", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "CoilRe:CreateUserDefinedPart:1", _
  "Bs2", "slot_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "CoilRe:CreateUserDefinedPart:1", _
  "LenRegion", "max(L, L) + 2*endRegion"
On Error Goto 0
oEditor.Rotate Array("NAME:Selections", "Selections:=", "CoilRe"), Array(_
  "NAME:RotateParameters", "CoordinateSystemID:=", -1, "RotateAxis:=", "Z", _
  "RotateAngle:=", "-11.428571428571429deg")
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "CoilRe"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", _
  "1.4285714285714286deg", "NumClones:=", "252"), Array("NAME:Options", _
  "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "CoilRe"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "CoilRe_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_0,CoilRe_63" & _
  ",CoilRe_126,CoilRe_189"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_1,CoilRe_64" & _
  ",CoilRe_127,CoilRe_190"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_2,CoilRe_65" & _
  ",CoilRe_128,CoilRe_191"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_3,CoilRe_66" & _
  ",CoilRe_129,CoilRe_192"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_4,CoilRe_67" & _
  ",CoilRe_130,CoilRe_193"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_5,CoilRe_68" & _
  ",CoilRe_131,CoilRe_194"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_6,CoilRe_69" & _
  ",CoilRe_132,CoilRe_195"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_7,CoilRe_70" & _
  ",CoilRe_133,CoilRe_196"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_8,CoilRe_71" & _
  ",CoilRe_134,CoilRe_197"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_9,CoilRe_72" & _
  ",CoilRe_135,CoilRe_198"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_10,CoilRe_73" & _
  ",CoilRe_136,CoilRe_199"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_11,CoilRe_74" & _
  ",CoilRe_137,CoilRe_200"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_12,CoilRe_75" & _
  ",CoilRe_138,CoilRe_201"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_13,CoilRe_76" & _
  ",CoilRe_139,CoilRe_202"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_14,CoilRe_77" & _
  ",CoilRe_140,CoilRe_203"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_15,CoilRe_78" & _
  ",CoilRe_141,CoilRe_204"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_16,CoilRe_79" & _
  ",CoilRe_142,CoilRe_205"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_17,CoilRe_80" & _
  ",CoilRe_143,CoilRe_206"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_18,CoilRe_81" & _
  ",CoilRe_144,CoilRe_207"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_19,CoilRe_82" & _
  ",CoilRe_145,CoilRe_208"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_20,CoilRe_83" & _
  ",CoilRe_146,CoilRe_209"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_21,CoilRe_84" & _
  ",CoilRe_147,CoilRe_210"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_22,CoilRe_85" & _
  ",CoilRe_148,CoilRe_211"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_23,CoilRe_86" & _
  ",CoilRe_149,CoilRe_212"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_24,CoilRe_87" & _
  ",CoilRe_150,CoilRe_213"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_25,CoilRe_88" & _
  ",CoilRe_151,CoilRe_214"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_26,CoilRe_89" & _
  ",CoilRe_152,CoilRe_215"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_27,CoilRe_90" & _
  ",CoilRe_153,CoilRe_216"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_28,CoilRe_91" & _
  ",CoilRe_154,CoilRe_217"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_29,CoilRe_92" & _
  ",CoilRe_155,CoilRe_218"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_30,CoilRe_93" & _
  ",CoilRe_156,CoilRe_219"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_31,CoilRe_94" & _
  ",CoilRe_157,CoilRe_220"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_32,CoilRe_95" & _
  ",CoilRe_158,CoilRe_221"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_33,CoilRe_96" & _
  ",CoilRe_159,CoilRe_222"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_34,CoilRe_97" & _
  ",CoilRe_160,CoilRe_223"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_35,CoilRe_98" & _
  ",CoilRe_161,CoilRe_224"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_36,CoilRe_99" & _
  ",CoilRe_162,CoilRe_225"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_37,CoilRe_100" & _
  ",CoilRe_163,CoilRe_226"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_38,CoilRe_101" & _
  ",CoilRe_164,CoilRe_227"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_39,CoilRe_102" & _
  ",CoilRe_165,CoilRe_228"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_40,CoilRe_103" & _
  ",CoilRe_166,CoilRe_229"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_41,CoilRe_104" & _
  ",CoilRe_167,CoilRe_230"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_42,CoilRe_105" & _
  ",CoilRe_168,CoilRe_231"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_43,CoilRe_106" & _
  ",CoilRe_169,CoilRe_232"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_44,CoilRe_107" & _
  ",CoilRe_170,CoilRe_233"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_45,CoilRe_108" & _
  ",CoilRe_171,CoilRe_234"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_46,CoilRe_109" & _
  ",CoilRe_172,CoilRe_235"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_47,CoilRe_110" & _
  ",CoilRe_173,CoilRe_236"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_48,CoilRe_111" & _
  ",CoilRe_174,CoilRe_237"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_49,CoilRe_112" & _
  ",CoilRe_175,CoilRe_238"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_50,CoilRe_113" & _
  ",CoilRe_176,CoilRe_239"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_51,CoilRe_114" & _
  ",CoilRe_177,CoilRe_240"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_52,CoilRe_115" & _
  ",CoilRe_178,CoilRe_241"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_53,CoilRe_116" & _
  ",CoilRe_179,CoilRe_242"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_54,CoilRe_117" & _
  ",CoilRe_180,CoilRe_243"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_55,CoilRe_118" & _
  ",CoilRe_181,CoilRe_244"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_56,CoilRe_119" & _
  ",CoilRe_182,CoilRe_245"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_57,CoilRe_120" & _
  ",CoilRe_183,CoilRe_246"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_58,CoilRe_121" & _
  ",CoilRe_184,CoilRe_247"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_59,CoilRe_122" & _
  ",CoilRe_185,CoilRe_248"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_60,CoilRe_123" & _
  ",CoilRe_186,CoilRe_249"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_61,CoilRe_124" & _
  ",CoilRe_187,CoilRe_250"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "CoilRe_62,CoilRe_125" & _
  ",CoilRe_188,CoilRe_251"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseA", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "11267.7*sin(2*pi*50*time+delta)", "Resistance:=", "R1", "Inductance:=", _
  "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseB", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "11267.7*sin(2*pi*50*time+delta-2*pi/3)", "Resistance:=", "R1", _
  "Inductance:=", "Le1", "ParallelBranchesNum:=", "1")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:PhaseC", "Type:=", "Voltage", _
  "IsSolid:=", false, "Current:=", "0A", "Voltage:=", _
  "11267.7*sin(2*pi*50*time+delta-4*pi/3)", "Resistance:=", "R1", _
  "Inductance:=", "Le1", "ParallelBranchesNum:=", "1")
oModule.AssignCoil Array("NAME:PhA_0", "Objects:=", Array("Coil_0"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_0", "Objects:=", Array("CoilRe_8"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_1", "Objects:=", Array("Coil_1"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_1", "Objects:=", Array("CoilRe_9"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_2", "Objects:=", Array("Coil_2"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_2", "Objects:=", Array("CoilRe_10"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_3", "Objects:=", Array("Coil_3"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_3", "Objects:=", Array("CoilRe_11"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_4", "Objects:=", Array("Coil_4"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_4", "Objects:=", Array("CoilRe_12"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_5", "Objects:=", Array("Coil_5"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_5", "Objects:=", Array("CoilRe_13"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_6", "Objects:=", Array("Coil_6"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_6", "Objects:=", Array("CoilRe_14"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_7", "Objects:=", Array("Coil_7"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_7", "Objects:=", Array("CoilRe_15"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_8", "Objects:=", Array("Coil_8"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_8", "Objects:=", Array("CoilRe_16"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_9", "Objects:=", Array("Coil_9"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_9", "Objects:=", Array("CoilRe_17"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_10", "Objects:=", Array("Coil_10"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_10", "Objects:=", Array("CoilRe_18"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_11", "Objects:=", Array("Coil_11"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_11", "Objects:=", Array("CoilRe_19"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_12", "Objects:=", Array("Coil_12"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_12", "Objects:=", Array("CoilRe_20"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_13", "Objects:=", Array("Coil_13"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_13", "Objects:=", Array("CoilRe_21"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_14", "Objects:=", Array("Coil_14"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_14", "Objects:=", Array("CoilRe_22"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_15", "Objects:=", Array("Coil_15"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_15", "Objects:=", Array("CoilRe_23"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_16", "Objects:=", Array("Coil_16"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_16", "Objects:=", Array("CoilRe_24"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_17", "Objects:=", Array("Coil_17"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_17", "Objects:=", Array("CoilRe_25"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_18", "Objects:=", Array("Coil_18"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_18", "Objects:=", Array("CoilRe_26"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_19", "Objects:=", Array("Coil_19"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_19", "Objects:=", Array("CoilRe_27"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_20", "Objects:=", Array("Coil_20"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_20", "Objects:=", Array("CoilRe_28"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_21", "Objects:=", Array("Coil_21"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_21", "Objects:=", Array("CoilRe_29"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_22", "Objects:=", Array("Coil_22"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_22", "Objects:=", Array("CoilRe_30"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_23", "Objects:=", Array("Coil_23"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_23", "Objects:=", Array("CoilRe_31"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_24", "Objects:=", Array("Coil_24"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_24", "Objects:=", Array("CoilRe_32"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_25", "Objects:=", Array("Coil_25"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_25", "Objects:=", Array("CoilRe_33"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_26", "Objects:=", Array("Coil_26"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_26", "Objects:=", Array("CoilRe_34"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_27", "Objects:=", Array("Coil_27"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_27", "Objects:=", Array("CoilRe_35"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_28", "Objects:=", Array("Coil_28"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_28", "Objects:=", Array("CoilRe_36"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_29", "Objects:=", Array("Coil_29"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_29", "Objects:=", Array("CoilRe_37"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_30", "Objects:=", Array("Coil_30"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_30", "Objects:=", Array("CoilRe_38"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_31", "Objects:=", Array("Coil_31"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_31", "Objects:=", Array("CoilRe_39"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_32", "Objects:=", Array("Coil_32"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_32", "Objects:=", Array("CoilRe_40"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_33", "Objects:=", Array("Coil_33"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_33", "Objects:=", Array("CoilRe_41"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_34", "Objects:=", Array("Coil_34"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_34", "Objects:=", Array("CoilRe_42"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_35", "Objects:=", Array("Coil_35"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_35", "Objects:=", Array("CoilRe_43"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_36", "Objects:=", Array("Coil_36"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_36", "Objects:=", Array("CoilRe_44"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_37", "Objects:=", Array("Coil_37"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_37", "Objects:=", Array("CoilRe_45"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_38", "Objects:=", Array("Coil_38"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_38", "Objects:=", Array("CoilRe_46"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_39", "Objects:=", Array("Coil_39"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_39", "Objects:=", Array("CoilRe_47"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_40", "Objects:=", Array("Coil_40"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_40", "Objects:=", Array("CoilRe_48"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_41", "Objects:=", Array("Coil_41"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_41", "Objects:=", Array("CoilRe_49"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_42", "Objects:=", Array("Coil_42"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_42", "Objects:=", Array("CoilRe_50"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_43", "Objects:=", Array("Coil_43"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_43", "Objects:=", Array("CoilRe_51"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_44", "Objects:=", Array("Coil_44"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_44", "Objects:=", Array("CoilRe_52"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_45", "Objects:=", Array("Coil_45"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_45", "Objects:=", Array("CoilRe_53"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_46", "Objects:=", Array("Coil_46"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_46", "Objects:=", Array("CoilRe_54"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_47", "Objects:=", Array("Coil_47"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_47", "Objects:=", Array("CoilRe_55"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_48", "Objects:=", Array("Coil_48"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_48", "Objects:=", Array("CoilRe_56"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_49", "Objects:=", Array("Coil_49"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_49", "Objects:=", Array("CoilRe_57"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_50", "Objects:=", Array("Coil_50"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_50", "Objects:=", Array("CoilRe_58"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_51", "Objects:=", Array("Coil_51"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_51", "Objects:=", Array("CoilRe_59"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_52", "Objects:=", Array("Coil_52"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_52", "Objects:=", Array("CoilRe_60"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_53", "Objects:=", Array("Coil_53"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_53", "Objects:=", Array("CoilRe_61"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_54", "Objects:=", Array("Coil_54"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_54", "Objects:=", Array("CoilRe_62"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_55", "Objects:=", Array("Coil_55"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhARe_56", "Objects:=", Array("Coil_56"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhARe_57", "Objects:=", Array("Coil_57"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhC_58", "Objects:=", Array("Coil_58"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_59", "Objects:=", Array("Coil_59"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhC_60", "Objects:=", Array("Coil_60"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhBRe_61", "Objects:=", Array("Coil_61"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_62", "Objects:=", Array("Coil_62"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhBRe_244", "Objects:=", Array("CoilRe_0"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhA_245", "Objects:=", Array("CoilRe_1"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhA_246", "Objects:=", Array("CoilRe_2"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseA")
oModule.AssignCoil Array("NAME:PhCRe_247", "Objects:=", Array("CoilRe_3"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_248", "Objects:=", Array("CoilRe_4"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhCRe_249", "Objects:=", Array("CoilRe_5"), _
  "Conductor number:=", "conds", "PolarityType:=", "Negative", "Winding:=", _
  "PhaseC")
oModule.AssignCoil Array("NAME:PhB_250", "Objects:=", Array("CoilRe_6"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oModule.AssignCoil Array("NAME:PhB_251", "Objects:=", Array("CoilRe_7"), _
  "Conductor number:=", "conds", "PolarityType:=", "Positive", "Winding:=", _
  "PhaseB")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5454.808326mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "450.3175995mm"), _
  Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", "98.78203737mm"), _
  Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "360.2540796mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "358.4315657mm"), _
  Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "0"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "7"))), _
  Array("NAME:Attributes", "Name:=", "Rotor", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "Cogent_2DSF0.950", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "DiaGap", "d_o"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Rotor:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
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
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "450.3175995mm"), _
  Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", "98.78203737mm"), _
  Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "360.2540796mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "358.4315657mm"), _
  Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "0"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "6"))), _
  Array("NAME:Attributes", "Name:=", "Pole", "Flags:=", "", "Color:=", _
  "(132 132 193)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "Cogent_2DSF0.970", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "DiaGap", "d_o"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Pole:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
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
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "450.3175995mm"), _
  Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", "98.78203737mm"), _
  Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "360.2540796mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "358.4315657mm"), _
  Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "0"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "4"))), _
  Array("NAME:Attributes", "Name:=", "Field", "Flags:=", "", "Color:=", _
  "(255 128 192)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "DiaGap", "d_o"
oEditor.SetPropertyValue "Geometry3DCmdTab", "Field:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
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
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_0,Field_8" & _
  ",Field_16,Field_24"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_1,Field_9" & _
  ",Field_17,Field_25"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_2,Field_10" & _
  ",Field_18,Field_26"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_3,Field_11" & _
  ",Field_19,Field_27"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_4,Field_12" & _
  ",Field_20,Field_28"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_5,Field_13" & _
  ",Field_21,Field_29"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_6,Field_14" & _
  ",Field_22,Field_30"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "Field_7,Field_15" & _
  ",Field_23,Field_31"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5454.808326mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "450.3175995mm"), _
  Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", "98.78203737mm"), _
  Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "360.2540796mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "358.4315657mm"), _
  Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "0"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", "5"))), _
  Array("NAME:Attributes", "Name:=", "FieldRe", "Flags:=", "", "Color:=", _
  "(255 128 192)", "Transparency:=", 0, "PartCoordinateSystem:=", "Global", _
  "MaterialName:=", "copper", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "DiaGap", "d_o"
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "DiaYoke", "d_i"
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "WidthShoe", "shoe_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "HeightShoe", "shoe_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "WidthBody", "pole_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "HeightBody", "pole_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", "FieldRe:CreateUserDefinedPart:1", _
  "LenRegion", "max(L, L) + 2*endRegion"
On Error Goto 0
oEditor.DuplicateAroundAxis Array("NAME:Selections", "Selections:=", "FieldRe"), _
  Array("NAME:DuplicateAroundAxisParameters", "CoordinateSystemID:=", -1, _
  "CreateNewObjects:=", true, "WhichAxis:=", "Z", "AngleStr:=", "11.25deg", _
  "NumClones:=", "32"), Array("NAME:Options", "DuplicateBoundaries:=", false)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "FieldRe"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "FieldRe_0"))))
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_0,FieldRe_8" & _
  ",FieldRe_16,FieldRe_24"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_1,FieldRe_9" & _
  ",FieldRe_17,FieldRe_25"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_2,FieldRe_10" & _
  ",FieldRe_18,FieldRe_26"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_3,FieldRe_11" & _
  ",FieldRe_19,FieldRe_27"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_4,FieldRe_12" & _
  ",FieldRe_20,FieldRe_28"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_5,FieldRe_13" & _
  ",FieldRe_21,FieldRe_29"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_6,FieldRe_14" & _
  ",FieldRe_22,FieldRe_30"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.Unite Array("NAME:Selections", "Selections:=", "FieldRe_7,FieldRe_15" & _
  ",FieldRe_23,FieldRe_31"), Array("NAME:UniteParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.AssignWindingGroup Array("NAME:Field", "Type:=", "Current", _
  "IsSolid:=", false, "Current:=", "340.809A", "Resistance:=", "0ohm", _
  "Inductance:=", "0H", "ParallelBranchesNum:=", "1")
oModule.AssignCoil Array("NAME:Field_0", "Objects:=", Array("Field_0"), _
  "Conductor number:=", 36, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_0", "Objects:=", Array("FieldRe_0"), _
  "Conductor number:=", 36, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_1", "Objects:=", Array("Field_1"), _
  "Conductor number:=", 36, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_1", "Objects:=", Array("FieldRe_1"), _
  "Conductor number:=", 36, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_2", "Objects:=", Array("Field_2"), _
  "Conductor number:=", 36, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_2", "Objects:=", Array("FieldRe_2"), _
  "Conductor number:=", 36, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_3", "Objects:=", Array("Field_3"), _
  "Conductor number:=", 36, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_3", "Objects:=", Array("FieldRe_3"), _
  "Conductor number:=", 36, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_4", "Objects:=", Array("Field_4"), _
  "Conductor number:=", 36, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_4", "Objects:=", Array("FieldRe_4"), _
  "Conductor number:=", 36, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_5", "Objects:=", Array("Field_5"), _
  "Conductor number:=", 36, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_5", "Objects:=", Array("FieldRe_5"), _
  "Conductor number:=", 36, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_6", "Objects:=", Array("Field_6"), _
  "Conductor number:=", 36, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_6", "Objects:=", Array("FieldRe_6"), _
  "Conductor number:=", 36, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:FieldRe_7", "Objects:=", Array("Field_7"), _
  "Conductor number:=", 36, "PolarityType:=", "Negative", "Winding:=", _
  "Field")
oModule.AssignCoil Array("NAME:Field_7", "Objects:=", Array("FieldRe_7"), _
  "Conductor number:=", 36, "PolarityType:=", "Positive", "Winding:=", _
  "Field")
oEditor.CreateUserDefinedPart Array("NAME:UserDefinedPrimitiveParameters", _
  "DllName:=", "RMxprt/SalientPoleCore", "Version:=", "12.1", _
  "NoOfParameters:=", 34, "Library:=", "syslib", Array("NAME:ParamVector", _
  Array("NAME:Pair", "Name:=", "DiaGap", "Value:=", "5454.808326mm"), Array(_
  "NAME:Pair", "Name:=", "DiaYoke", "Value:=", "4000mm"), Array("NAME:Pair", _
  "Name:=", "Length", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Skew", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Slots", "Value:=", "0"), _
  Array("NAME:Pair", "Name:=", "SlotType", "Value:=", "0"), Array("NAME:Pair", _
  "Name:=", "Hs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs01", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Hs1", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Hs2", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Bs0", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs1", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "Bs2", "Value:=", "0mm"), _
  Array("NAME:Pair", "Name:=", "Rs", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "FilletType", "Value:=", "0"), Array("NAME:Pair", "Name:=", _
  "SlotPitch", "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "CenterPitch", _
  "Value:=", "0deg"), Array("NAME:Pair", "Name:=", "Poles", "Value:=", "32"), _
  Array("NAME:Pair", "Name:=", "WidthShoe", "Value:=", "450.3175995mm"), _
  Array("NAME:Pair", "Name:=", "HeightShoe", "Value:=", "98.78203737mm"), _
  Array("NAME:Pair", "Name:=", "WidthBody", "Value:=", "360.2540796mm"), _
  Array("NAME:Pair", "Name:=", "HeightBody", "Value:=", "358.4315657mm"), _
  Array("NAME:Pair", "Name:=", "AirGap2", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "Offset", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "Off2_x", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "Off2_y", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "CoilEndExt", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "EndRingType", "Value:=", _
  "0"), Array("NAME:Pair", "Name:=", "BarEndExt", "Value:=", "0mm"), Array(_
  "NAME:Pair", "Name:=", "RingLength", "Value:=", "0mm"), Array("NAME:Pair", _
  "Name:=", "RingHeight", "Value:=", "0mm"), Array("NAME:Pair", "Name:=", _
  "SegAngle", "Value:=", "15deg"), Array("NAME:Pair", "Name:=", "LenRegion", _
  "Value:=", "0mm"), Array("NAME:Pair", "Name:=", "InfoCore", "Value:=", _
  "100"))), Array("NAME:Attributes", "Name:=", "InnerRegion", "Flags:=", "", _
  "Color:=", "(0 255 255)", "Transparency:=", 0, "PartCoordinateSystem:=", _
  "Global", "MaterialName:=", "vacuum", "SolveInside:=", true) 
On Error Resume Next
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "DiaGap", "d_o"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "DiaYoke", "d_i"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "WidthShoe", "shoe_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "HeightShoe", "shoe_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "WidthBody", "pole_w"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "HeightBody", "pole_h"
oEditor.SetPropertyValue "Geometry3DCmdTab", _
  "InnerRegion:CreateUserDefinedPart:1", "LenRegion", _
  "max(L, L) + 2*endRegion"
On Error Goto 0
Set oModule = oDesign.GetModule("MeshSetup")
oModule.AssignTrueSurfOp Array("NAME:SurfApprox_Main", "Objects:=", Array(_
  "Stator", "Rotor", "Band", "OuterRegion", "InnerRegion", "Shaft"), _
  "SurfDevChoice:=", 2, "SurfDev:=", "3.04721mm", "NormalDevChoice:=", 2, _
  "NormalDev:=", "15deg", "AspectRatioChoice:=", 1)
oEditor.ChangeProperty Array("NAME:AllTabs", Array(_
  "NAME:Geometry3DAttributeTab", Array("NAME:PropServers", "Band", _
  "OuterRegion", "InnerRegion"), Array("NAME:ChangedProps", Array(_
  "NAME:Transparent", "Value:=", 0.75))))
oEditor.Subtract Array("NAME:Selections", "Blank Parts:=", "Band,InnerRegion" & _
  ",Shaft,Stator,Coil_0,Coil_1,Coil_2,Coil_3,Coil_4,Coil_5,Coil_6,Coil_7" & _
  ",Coil_8,Coil_9,Coil_10,Coil_11,Coil_12,Coil_13,Coil_14,Coil_15,Coil_16" & _
  ",Coil_17,Coil_18,Coil_19,Coil_20,Coil_21,Coil_22,Coil_23,Coil_24" & _
  ",Coil_25,Coil_26,Coil_27,Coil_28,Coil_29,Coil_30,Coil_31,Coil_32" & _
  ",Coil_33,Coil_34,Coil_35,Coil_36,Coil_37,Coil_38,Coil_39,Coil_40" & _
  ",Coil_41,Coil_42,Coil_43,Coil_44,Coil_45,Coil_46,Coil_47,Coil_48" & _
  ",Coil_49,Coil_50,Coil_51,Coil_52,Coil_53,Coil_54,Coil_55,Coil_56" & _
  ",Coil_57,Coil_58,Coil_59,Coil_60,Coil_61,Coil_62,CoilRe_0,CoilRe_1" & _
  ",CoilRe_2,CoilRe_3,CoilRe_4,CoilRe_5,CoilRe_6,CoilRe_7,CoilRe_8" & _
  ",CoilRe_9,CoilRe_10,CoilRe_11,CoilRe_12,CoilRe_13,CoilRe_14,CoilRe_15" & _
  ",CoilRe_16,CoilRe_17,CoilRe_18,CoilRe_19,CoilRe_20,CoilRe_21,CoilRe_22" & _
  ",CoilRe_23,CoilRe_24,CoilRe_25,CoilRe_26,CoilRe_27,CoilRe_28,CoilRe_29" & _
  ",CoilRe_30,CoilRe_31,CoilRe_32,CoilRe_33,CoilRe_34,CoilRe_35,CoilRe_36" & _
  ",CoilRe_37,CoilRe_38,CoilRe_39,CoilRe_40,CoilRe_41,CoilRe_42,CoilRe_43" & _
  ",CoilRe_44,CoilRe_45,CoilRe_46,CoilRe_47,CoilRe_48,CoilRe_49,CoilRe_50" & _
  ",CoilRe_51,CoilRe_52,CoilRe_53,CoilRe_54,CoilRe_55,CoilRe_56,CoilRe_57" & _
  ",CoilRe_58,CoilRe_59,CoilRe_60,CoilRe_61,CoilRe_62,Rotor,Pole,Field_0" & _
  ",Field_1,Field_2,Field_3,Field_4,Field_5,Field_6,Field_7,FieldRe_0" & _
  ",FieldRe_1,FieldRe_2,FieldRe_3,FieldRe_4,FieldRe_5,FieldRe_6,FieldRe_7", _
  "Tool Parts:=", "Tool"), Array("NAME:SubtractParameters", _
  "CoordinateSystemID:=", -1, "KeepOriginals:=", false)
oEditor.FitAll
Set oModule = oDesign.GetModule("ModelSetup")
oModule.AssignBand Array("NAME:MotionSetup1", "Move Type:=", "Rotate", _
  "Coordinate System:=", "Global", "Axis:=", "Z", "Is Positive:=", true, _
  "InitPos:=", "13.8393deg", "HasRotateLimit:=", false, "NonCylindrical:=", _
  false, "Consider Mechanical Transient:=", true, "Angular Velocity:=", _
  "187.5rpm", "Moment of Inertia:=", 730158, "Damping:=", 794.053, _
  "Load Torque:=", _
  "if(speed<19.635, -100702*speed, -3.88236e+07/speed) - 100702*(speed-19.635)*10", _
  "Objects:=", Array("Band"))
oModule.EditMotionSetup "MotionSetup1", Array("NAME:Data", _
  "Consider Mechanical Transient:=", false)
Set oModule = oDesign.GetModule("AnalysisSetup")
oModule.InsertSetup "Transient", Array("NAME:Setup1", "FastReachSteadyState:=", _
  true, "AutoDetectSteadyState:=", true, "FrequencyOfAddedVoltageSource:=", _
  "50Hz", "StopTime:=", "0.2s", "TimeStep:=", "0.0002s")
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetMinimumTimeStep "0.002ms" 
Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "Torque", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Moving1.Torque")), Array()
oModule.CreateReport "Currents", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Current(PhaseA)", "Current(PhaseB)", _
  "Current(PhaseC)")), Array()
oModule.CreateReport "Induced Voltages", "Transient", "XY Plot", _
  "Setup1 : Transient", Array(), Array("Time:=", Array("All")), Array(_
  "X Component:=", "Time", "Y Component:=", Array("InducedVoltage(PhaseA)", _
  "InducedVoltage(PhaseB)", "InducedVoltage(PhaseC)")), Array()
oModule.CreateReport "Flux Linkages", "Transient", "XY Plot", _
  "Setup1 : Transient", Array(), Array("Time:=", Array("All")), Array(_
  "X Component:=", "Time", "Y Component:=", Array("FluxLinkage(PhaseA)", _
  "FluxLinkage(PhaseB)", "FluxLinkage(PhaseC)")), Array()
oModule.CreateReport "Voltages", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("InputVoltage(PhaseA)", "InputVoltage(PhaseB)", _
  "InputVoltage(PhaseC)")), Array()
oModule.CreateReport "Powers", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array(_
  "InputVoltage(PhaseA)*Current(PhaseA) + InputVoltage(PhaseB)*Current(PhaseB) + InputVoltage(PhaseC)*Current(PhaseC)", _
  "Moving1.Speed*Moving1.Torque")), Array()
oModule.ChangeProperty Array("NAME:AllTabs", Array("NAME:Trace", Array(_
  "NAME:PropServers", _
  "Powers:InputVoltage(PhaseA)*Current(PhaseA) + InputVoltage(PhaseB)*Current(PhaseB) + InputVoltage(PhaseC)*Current(PhaseC)"), _
  Array("NAME:ChangedProps", Array("NAME:Specify Name", "Value:=", true))))
oModule.ChangeProperty Array("NAME:AllTabs", Array("NAME:Trace", Array(_
  "NAME:PropServers", _
  "Powers:InputVoltage(PhaseA)*Current(PhaseA) + InputVoltage(PhaseB)*Current(PhaseB) + InputVoltage(PhaseC)*Current(PhaseC)"), _
  Array("NAME:ChangedProps", Array("NAME:Name", "Value:=", "ElecPower"))))
oModule.ChangeProperty Array("NAME:AllTabs", Array("NAME:Trace", Array(_
  "NAME:PropServers", "Powers:Moving1.Speed*Moving1.Torque"), Array(_
  "NAME:ChangedProps", Array("NAME:Specify Name", "Value:=", true))))
oModule.ChangeProperty Array("NAME:AllTabs", Array("NAME:Trace", Array(_
  "NAME:PropServers", "Powers:Moving1.Speed*Moving1.Torque"), Array(_
  "NAME:ChangedProps", Array("NAME:Name", "Value:=", "MechPower"))))
oModule.AddTraceCharacteristics "Torque", "avg", Array(), Array("Specified", _
  "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Currents", "rms", Array(), Array("Specified", _
  "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Currents", "avg", Array(), Array("Specified", _
  "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Voltages", "rms", Array(), Array("Specified", _
  "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Induced Voltages", "rms", Array(), Array(_
  "Specified", "0.02s", "0.04s")
oModule.AddTraceCharacteristics "Powers", "avg", Array(), Array("Specified", _
  "0.02s", "0.04s")
oEditor.ShowWindow
Set oModule = oDesign.GetModule("OutputVariable")
oModule.CreateOutputVariable "pos", _
  "(Moving1.Position -13.8393 * PI/180) * 16 + PI", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos0", "cos(pos)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos1", "cos(pos-2*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "cos2", "cos(pos-4*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin0", "sin(pos)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin1", "sin(pos-2*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "sin2", "sin(pos-4*PI/3)", "Setup1 : Transient", _
  "Transient", Array() 
oModule.CreateOutputVariable "Lad", _
  "L(PhaseA,PhaseA)*cos0 + L(PhaseA,PhaseB)*cos1 + L(PhaseA,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Laq", _
  "L(PhaseA,PhaseA)*sin0 + L(PhaseA,PhaseB)*sin1 + L(PhaseA,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lbd", _
  "L(PhaseB,PhaseA)*cos0 + L(PhaseB,PhaseB)*cos1 + L(PhaseB,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lbq", _
  "L(PhaseB,PhaseA)*sin0 + L(PhaseB,PhaseB)*sin1 + L(PhaseB,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lcd", _
  "L(PhaseC,PhaseA)*cos0 + L(PhaseC,PhaseB)*cos1 + L(PhaseC,PhaseC)*cos2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Lcq", _
  "L(PhaseC,PhaseA)*sin0 + L(PhaseC,PhaseB)*sin1 + L(PhaseC,PhaseC)*sin2", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "L_d", "(Lad*cos0 + Lbd*cos1 + Lcd*cos2) * 2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "L_q", "(Laq*sin0 + Lbq*sin1 + Lcq*sin2) * 2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Flux_d", _
  "(FluxLinkage(PhaseA)*cos0+FluxLinkage(PhaseB)*cos1+FluxLinkage(PhaseC)*cos2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Flux_q", _
  "(FluxLinkage(PhaseA)*sin0+FluxLinkage(PhaseB)*sin1+FluxLinkage(PhaseC)*sin2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "I_d", _
  "(Current(PhaseA)*cos0 + Current(PhaseB)*cos1 + Current(PhaseC)*cos2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "I_q", _
  "(Current(PhaseA)*sin0 + Current(PhaseB)*sin1 + Current(PhaseC)*sin2)*2/3", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Irms", "sqrt(I_d^2+I_q^2)/sqrt(2)", _
  "Setup1 : Transient", "Transient", Array() 
oModule.CreateOutputVariable "Pcu", "3*Irms^2*0", "Setup1 : Transient", _
  "Transient", Array() 
Set oModule = oDesign.GetModule("ReportSetup")
oModule.CreateReport "L_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("L_d", "L_q")), Array()
oModule.CreateReport "Flux_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("Flux_d", "Flux_q")), Array()
oModule.CreateReport "I_dq", "Transient", "XY Plot", "Setup1 : Transient", _
  Array(), Array("Time:=", Array("All")), Array("X Component:=", "Time", _
  "Y Component:=", Array("I_d", "I_q")), Array()
oDesign.SetDesignSettings Array("NAME:Design Settings Data", _
  "ComputeTransientInductance:=", true, "ComputeIncrementalMatrix:=", false)
Set oModule = oDesign.GetModule("BoundarySetup")
oModule.SetupYConnection Array(Array("NAME:YConnection", "Windings:=", _
  "PhaseA,PhaseB,PhaseC"))
