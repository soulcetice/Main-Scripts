
'This script is used to display important data for each skid (Diameter,Width,Coil Id) in the Coil Car Overview pictures
Dim MM	
Dim No
Dim i, j
Dim Arr(25)
Dim Skid(25)
Dim Table(25)

Dim group 
Set group = HMIRuntime.Tags.CreateTagSet

'You must define Tracking Position for each Coil
	group.Add "SMS_TCM::MM1_HmiPieceDspAddDat1_TrPos"
	group.Add "SMS_TCM::MM1_HmiPieceDspAddDat2_TrPos"
	group.Add "SMS_TCM::MM1_HmiPieceDspAddDat3_TrPos"
	group.Add "SMS_TCM::MM1_HmiPieceDspAddDat4_TrPos"
	group.Add "SMS_TCM::MM1_HmiPieceDspAddDat5_TrPos"
	
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat1_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat2_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat3_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat4_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat5_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat6_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat7_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat8_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat9_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat10_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat11_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat12_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat13_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat14_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat15_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat16_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat17_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat18_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat19_TrPos"
	group.Add "SMS_TCM::MM2_HmiPieceDspAddDat20_TrPos"
	group.Read

	Arr(1) = group("SMS_TCM::MM2_HmiPieceDspAddDat1_TrPos").Value
	Arr(2) = group("SMS_TCM::MM2_HmiPieceDspAddDat2_TrPos").Value
	Arr(3) = group("SMS_TCM::MM2_HmiPieceDspAddDat3_TrPos").Value
	Arr(4) = group("SMS_TCM::MM2_HmiPieceDspAddDat4_TrPos").Value
	Arr(5) = group("SMS_TCM::MM2_HmiPieceDspAddDat5_TrPos").Value
	Arr(6) = group("SMS_TCM::MM2_HmiPieceDspAddDat6_TrPos").Value
	Arr(7) = group("SMS_TCM::MM2_HmiPieceDspAddDat7_TrPos").Value
	Arr(8) = group("SMS_TCM::MM2_HmiPieceDspAddDat8_TrPos").Value
	Arr(9) = group("SMS_TCM::MM2_HmiPieceDspAddDat9_TrPos").Value
	Arr(10) = group("SMS_TCM::MM2_HmiPieceDspAddDat10_TrPos").Value
	Arr(11) = group("SMS_TCM::MM2_HmiPieceDspAddDat11_TrPos").Value
	Arr(12) = group("SMS_TCM::MM2_HmiPieceDspAddDat12_TrPos").Value
	Arr(13) = group("SMS_TCM::MM2_HmiPieceDspAddDat13_TrPos").Value
	Arr(14) = group("SMS_TCM::MM2_HmiPieceDspAddDat14_TrPos").Value
	Arr(15) = group("SMS_TCM::MM2_HmiPieceDspAddDat15_TrPos").Value
	Arr(16) = group("SMS_TCM::MM2_HmiPieceDspAddDat16_TrPos").Value
	Arr(17) = group("SMS_TCM::MM2_HmiPieceDspAddDat17_TrPos").Value
	Arr(18) = group("SMS_TCM::MM2_HmiPieceDspAddDat18_TrPos").Value
	Arr(19) = group("SMS_TCM::MM2_HmiPieceDspAddDat19_TrPos").Value
	Arr(20) = group("SMS_TCM::MM2_HmiPieceDspAddDat20_TrPos").Value
	Arr(21) = group("SMS_TCM::MM1_HmiPieceDspAddDat1_TrPos").Value
	Arr(22) = group("SMS_TCM::MM1_HmiPieceDspAddDat2_TrPos").Value
	Arr(23) = group("SMS_TCM::MM1_HmiPieceDspAddDat3_TrPos").Value
	Arr(24) = group("SMS_TCM::MM1_HmiPieceDspAddDat4_TrPos").Value
	Arr(25) = group("SMS_TCM::MM1_HmiPieceDspAddDat5_TrPos").Value


	MM = Array(0,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1,1,1)
	No = Array(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,1,2,3,4,5)
	
'Take each piece, look at the position of it and store the data from that piece on the internal tags created for each position
	For i = 1 To 25 Step 1
		Skid(i) = 0
	Next
	
j=21	
    
    Dim groupWrite
    Set groupWrite = HMIRuntime.Tags.CreateTagSet
	Dim groupRead
	Set groupRead = HMIRuntime.Tags.CreateTagSet
	For i = 1 To 25 Step 1
		If  HMIRuntime.Tags("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_PieceStatInf").Read <> 0 Then
			If Arr(i) = 0 Then
				on error resume next
				groupRead.RemoveAll
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_InPieceId"
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_OutPieceId"
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Alloy"
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cid"
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cod"
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_StrLen"
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_StrThk"
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_CoilWght"
				on error resume next
                groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cwd"
                groupRead.Read

                groupWrite.Add "iCoilIdSkid_" & j 	
                groupWrite.Add "iCoilOutIdSkid_" & j
                groupWrite.Add "iAlloySkid_"  & j 	
                groupWrite.Add "iInnDiaSkid_" & j 	
                groupWrite.Add "iOutDiaSkid_" & j 	
                groupWrite.Add "iLenSkid_"    & j 	
                groupWrite.Add "iThickSkid_"  & j 	
                groupWrite.Add "iWeightSkid_" & j 	
                groupWrite.Add "iWidthSkid_"  & j 	
                groupWrite.Add "iPieceSkid_"  & j 	
                groupWrite.Add "iTrPosSkid_"  & j

                groupWrite("iCoilIdSkid_" & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_InPieceId").Value
                groupWrite("iCoilOutIdSkid_" & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_OutPieceId").Value
                groupWrite("iAlloySkid_"  & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Alloy").Value
                groupWrite("iInnDiaSkid_" & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cid").Value
                groupWrite("iOutDiaSkid_" & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cod").Value
                groupWrite("iLenSkid_"    & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_StrLen").Value
                groupWrite("iThickSkid_"  & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_StrThk").Value
                groupWrite("iWeightSkid_" & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_CoilWght").Value
                groupWrite("iWidthSkid_"  & j).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cwd").Value
                groupWrite("iPieceSkid_"  & j).Value = No(i)
                groupWrite("iTrPosSkid_"  & j).Value = "PL"
                
                Skid(j) = 1
                j = j + 1
			Else
				on error resume next
				groupRead.RemoveAll
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_InPieceId"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_OutPieceId"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Alloy"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cid"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cod"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_StrLen"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_StrThk"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_CoilWght"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cwd"
				on error resume next
				groupRead.Add "SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos"
				groupRead.Read

				groupWrite.Add "iCoilIdSkid_"      & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iCoilOutIdSkid_"   & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iAlloySkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iInnDiaSkid_"      & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iOutDiaSkid_"      & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iLenSkid_"         & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iThickSkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iWeightSkid_"      & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iWidthSkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iTrPosSkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value
				groupWrite.Add "iPieceSkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value

				groupWrite("iCoilIdSkid_"      & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_InPieceId").Value
				groupWrite("iCoilOutIdSkid_"   & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_OutPieceId").Value
				groupWrite("iAlloySkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Alloy").Value
				groupWrite("iInnDiaSkid_"      & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cid").Value
				groupWrite("iOutDiaSkid_"      & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cod").Value
				groupWrite("iLenSkid_"         & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_StrLen").Value
				groupWrite("iThickSkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_StrThk").Value
				groupWrite("iWeightSkid_"      & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_CoilWght").Value
				groupWrite("iWidthSkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_Cwd").Value
				groupWrite("iTrPosSkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = groupRead("SMS_TCM::MM" & MM(i) & "_HmiPieceDspGenDat" & No(i) & "_TrPos").Value
				groupWrite("iPieceSkid_"       & group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value).Value = No(i)	

                Skid(group("SMS_TCM::MM" & MM(i) & "_HmiPieceDspAddDat" & No(i) & "_TrPos").Value) = 1
		End If
		End If 
	Next

groupWrite.Write
groupWrite.RemoveAll

'Reset the data on the internal variables when the Coil leaves the Skid
	For i = 1 To 25 Step 1
		If Skid(i) = 0 Then		
			
			on error resume next
			groupWrite.Add "iCoilIdSkid_" & i
			on error resume next
			groupWrite.Add "iCoilOutIdSkid_" & i
			on error resume next
			groupWrite.Add "iAlloySkid_"  & i
			on error resume next
			groupWrite.Add "iInnDiaSkid_" & i
			on error resume next
			groupWrite.Add "iOutDiaSkid_" & i
			on error resume next
			groupWrite.Add "iLenSkid_" & i
			on error resume next
			groupWrite.Add "iThickSkid_"  & i
			on error resume next
			groupWrite.Add "iWeightSkid_" & i
			on error resume next
			groupWrite.Add "iWidthSkid_"  & i
			on error resume next
			groupWrite.Add "iPieceSkid_"  & i

			groupWrite("iCoilIdSkid_" & i).Value =  	""
			groupWrite("iCoilOutIdSkid_" & i).Value =  	""
			groupWrite("iAlloySkid_"  & i).Value =  	""
			groupWrite("iInnDiaSkid_" & i).Value =  	0
			groupWrite("iOutDiaSkid_" & i).Value =  	0
			groupWrite("iLenSkid_"    & i).Value =  	0
			groupWrite("iThickSkid_"  & i).Value =  	0
			groupWrite("iWeightSkid_" & i).Value =  	0
			groupWrite("iWidthSkid_"  & i).Value =  	0
			groupWrite("iPieceSkid_"  & i).Value =  	0
		End If 
	Next

'Overview Table Section

	For i = 1 To 25 Step 1
		Table(i) = 0
	Next
	
groupWrite.Write
groupWrite.RemoveAll

j=1	
	For i = 1 To 25 Step 1
		If  HMIRuntime.Tags("iWidthSkid_" & i).Read <> 0 Then

			On Error Resume Next
			groupRead.RemoveAll			
			groupRead.Add "iCoilIdSkid_" & i
			groupRead.Add "iCoilOutIdSkid_" & i
			groupRead.Add "iAlloySkid_"  & i
			groupRead.Add "iInnDiaSkid_" & i
			groupRead.Add "iOutDiaSkid_" & i
			groupRead.Add "iLenSkid_" & i
			groupRead.Add "iThickSkid_"  & i
			groupRead.Add "iWeightSkid_" & i
			groupRead.Add "iWidthSkid_"  & i
			groupRead.Add "iPieceSkid_"  & i
			groupRead.Add "iTrPosSkid_"  & i
			groupRead.Read

			groupWrite.Add "iCoilIdTable_" & i
			groupWrite.Add "iCoilOutIdTable_" & i
			groupWrite.Add "iAlloyTable_"  & i
			groupWrite.Add "iInnDiaTable_" & i
			groupWrite.Add "iOutDiaTable_" & i
			groupWrite.Add "iLenTable_" & i
			groupWrite.Add "iThickTable_"  & i
			groupWrite.Add "iWeightTable_" & i
			groupWrite.Add "iWidthTable_"  & i
			groupWrite.Add "iPieceTable_"  & i
			groupWrite.Add "iTrPosTable_"  & i

			groupWrite("iCoilIdTable_" & j).Value 			= groupRead("iCoilIdSkid_" & i).Value
			groupWrite("iCoilOutIdTable_" & j).Value 		= groupRead("iCoilOutIdSkid_" & i).Value
			groupWrite("iAlloyTable_"  & j).Value 			= groupRead("iAlloySkid_" & i).Value
			groupWrite("iInnDiaTable_" & j).Value 			= groupRead("iInnDiaSkid_" & i).Value
			groupWrite("iOutDiaTable_" & j).Value 			= groupRead("iOutDiaSkid_" & i).Value
			groupWrite("iLenTable_"    & j).Value 			= groupRead("iLenSkid_" & i).Value
			groupWrite("iThickTable_"  & j).Value 			= groupRead("iThickSkid_" & i).Value
			groupWrite("iWeightTable_" & j).Value 			= groupRead("iWeightSkid_" & i).Value
			groupWrite("iWidthTable_"  & j).Value 			= groupRead("iWidthSkid_" & i).Value
			groupWrite("iPieceTable_"  & j).Value 			= groupRead("iPieceSkid_" & i).Value
			groupWrite("iTrPosTable_"  & j).Value 			= groupRead("iTrPosSkid_" & i).Value

			Table(j) = 1
			j = j + 1
		Else		
			On Error Resume Next
			groupWrite.Add "iCoilIdTable_" & i
			On Error Resume Next
			groupWrite.Add "iCoilOutIdTable_" & i
			On Error Resume Next
			groupWrite.Add "iAlloyTable_"  & i
			On Error Resume Next
			groupWrite.Add "iInnDiaTable_" & i
			On Error Resume Next
			groupWrite.Add "iOutDiaTable_" & i
			On Error Resume Next
			groupWrite.Add "iLenTable_" & i
			On Error Resume Next
			groupWrite.Add "iThickTable_"  & i
			On Error Resume Next
			groupWrite.Add "iWeightTable_" & i
			On Error Resume Next
			groupWrite.Add "iWidthTable_"  & i
			On Error Resume Next
			groupWrite.Add "iPieceTable_"  & i
			On Error Resume Next
			groupWrite.Add "iTrPosTable_"  & i

			groupWrite("iCoilIdTable_" & i).Value = ""
			groupWrite("iCoilOutIdTable_" & i).Value = ""
			groupWrite("iAlloyTable_"  & i).Value = ""
			groupWrite("iInnDiaTable_" & i).Value = 0
			groupWrite("iOutDiaTable_" & i).Value = 0
			groupWrite("iLenTable_"    & i).Value = 0
			groupWrite("iThickTable_"  & i).Value = 0
			groupWrite("iWeightTable_" & i).Value = 0
			groupWrite("iWidthTable_"  & i).Value = 0
			groupWrite("iPieceTable_"  & i).Value = 0
			groupWrite("iTrPosTable_"  & i).Value = ""
		End If
	Next

groupWrite.Write
groupWrite.RemoveAll