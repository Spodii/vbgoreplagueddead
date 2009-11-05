Attribute VB_Name = "Skills"
Option Explicit

'Class requirements
' * Civilian = 1
' * Reaver = 2
' * Engineer = 4
' * Infiltrator = 8
' * SquadLeader = 16
' * Job = 32
' * None = -1

'Reaver (ID = 2)
Private Const Rush_ClassReq As Integer = 2
Private Const Whirlwind_ClassReq As Integer = 2
Private Const Bash_ClassReq As Integer = 2
Private Const Warcry_ClassReq As Integer = 2
Private Const Charge_ClassReq As Integer = 2
Private Const Grab_ClassReq As Integer = 2
Private Const CrackArmor_ClassReq As Integer = 2
Private Const EF_ClassReq As Integer = 2
Private Const Berserk_ClassReq As Integer = 2
Private Const RageExplosion_ClassReq As Integer = 2
Private Const Throw_ClassReq As Integer = 2

'Engineer (ID = 4)

'Infiltrator (ID = 8)
Private Const Hide_ClassReq As Integer = 8

'SquadLeader (ID = 16)

'General
Private Const MaxSummons As Byte = 3        'Maximum number of characters on player can summon

Public Sub Skill_Berserk_PC(ByVal CasterIndex As Integer)

'*****************************************************************
'Reaver casts berserk on self
'*****************************************************************
    
    'Ready the buffer for the icon
    ConBuf.PreAllocate 4

    'Change the berserk status
    If UserList(CasterIndex).Flags.Berserk Then
        UserList(CasterIndex).Flags.Berserk = 0
        ConBuf.Put_Byte DataCode.Server_RemoveIcon
    Else
        UserList(CasterIndex).Flags.Berserk = 1
        ConBuf.Put_Byte DataCode.Server_SetIcon
    End If
    
    'Send the berserk icon
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte IconID.Berserk
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer()
    
    'Update the user's modstats
    User_UpdateModStats CasterIndex

End Sub

Public Sub Skill_Hide_PC(ByVal CasterIndex As Integer)

'*****************************************************************
'Infiltrator hides
'*****************************************************************

    'Check if already hiding
    If UserList(CasterIndex).Flags.Hiding = 1 Then Exit Sub

    'Set the using as hiding
    UserList(CasterIndex).Flags.Hiding = 1
    
    'Send the packet to everyone in the map that the user is now hiding
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.User_Hide
    ConBuf.Put_Integer UserList(CasterIndex).Char.CharIndex
    ConBuf.Put_Byte 1
    Data_Send ToMap, 0, ConBuf.Get_Buffer(), UserList(CasterIndex).Pos.Map
    
    'Update the modstats
    User_UpdateModStats CasterIndex

End Sub

Public Sub Skill_CrackArmor_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Reaver cracks the armor of a NPC
'*****************************************************************

    'Make sure the NPC is valid
    If NPCList(TargetIndex).Flags.NPCAlive = 0 Then Exit Sub
    If NPCList(TargetIndex).Flags.NPCActive = 0 Then Exit Sub
    If NPCList(TargetIndex).Attackable = 0 Then Exit Sub
    
    'Check if the NPC's armor was already and is going to last at least 3 seconds more
    If NPCList(TargetIndex).ActiveSkills.CrackArmorTime > 0 Then
        If NPCList(TargetIndex).ActiveSkills.CrackArmorTime + 3000 < timeGetTime Then Exit Sub
    End If
    
    'Crack the NPC's armor
    NPCList(TargetIndex).ActiveSkills.CrackArmorTime = timeGetTime + 20000  '20 seconds
    
    'Reduce the user's EP
    UserList(CasterIndex).Stats.BaseStat(SID.MinEP) = UserList(CasterIndex).Stats.BaseStat(SID.MinEP) - 5
    
    'Update the NPC's mod stats
    NPC_UpdateModStats TargetIndex
    
    'Create the icon on the NPC
    ConBuf.PreAllocate 4
    ConBuf.Put_Byte DataCode.Server_SetIcon
    ConBuf.Put_Integer NPCList(TargetIndex).Char.CharIndex
    ConBuf.Put_Byte IconID.CrackArmor
    Data_Send ToMap, 0, ConBuf.Get_Buffer(), NPCList(TargetIndex).Pos.Map, PP_StatusIcons

End Sub

Public Sub Skill_Grab_PCtoNPC_Release(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Ends a grabbing process
'*****************************************************************

    'Remove grabbing for the caster
    If CasterIndex > 0 Then
        If CasterIndex <= LastUser Then
            UserList(CasterIndex).Flags.GrabChar = 0
        End If
    End If
    
    'Remove grabbing for the target
    If TargetIndex > 0 Then
        If TargetIndex <= LastNPC Then
            NPCList(TargetIndex).Flags.GrabbedBy = 0
        End If
    End If
    
    'Update the user and NPC's mod stats
    User_UpdateModStats CasterIndex
    NPC_UpdateModStats TargetIndex

End Sub

Public Sub Skill_Grab_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Reaver grabs their target, preventing them from moving
'*****************************************************************

    'Make sure the NPC is valid
    If NPCList(TargetIndex).Flags.NPCAlive = 0 Then Exit Sub
    If NPCList(TargetIndex).Flags.NPCActive = 0 Then Exit Sub
    If NPCList(TargetIndex).Attackable = 0 Then Exit Sub
    
    'Make sure the user isn't already grabbing something or the NPC isn't already grabbed
    If UserList(CasterIndex).Flags.GrabChar > 0 Then Exit Sub
    If NPCList(TargetIndex).Flags.GrabbedBy > 0 Then Exit Sub
    
    'Grab the NPC
    UserList(CasterIndex).Flags.GrabChar = NPCList(TargetIndex).Char.CharIndex
    NPCList(TargetIndex).Flags.GrabbedBy = UserList(CasterIndex).Char.CharIndex
    
    'Update the user and NPC's mod stats
    User_UpdateModStats CasterIndex
    NPC_UpdateModStats TargetIndex
    
    '//!!
    ConBuf.PreAllocate 2
    ConBuf.Put_Byte DataCode.Comm_Talk
    ConBuf.Put_String "You grab the " & NPCList(TargetIndex).Name & "."
    ConBuf.Put_Byte DataCode.Comm_FontType_Info
    Data_Send ToIndex, CasterIndex, ConBuf.Get_Buffer
    
End Sub

Public Sub Skill_Charge_PC(ByVal CasterIndex As Integer)

'*****************************************************************
'Reaver charges in the direction they're facing, pushing all in the way
'*****************************************************************
Const MoveTiles As Integer = 5
Dim NPCIndex As Integer
Dim PushDir(1 To 2) As Byte
Dim HitTilesX(1 To MoveTiles) As Integer
Dim HitTilesY(1 To MoveTiles) As Integer
Dim OldPos As WorldPos
Dim ToX As Integer
Dim ToY As Integer
Dim X As Integer
Dim Y As Integer
Dim i As Long

    With UserList(CasterIndex)
    
        'Store the user's old position
        OldPos = .Pos
    
        'Find the position the user is heading to, the tiles they will hit and the directions they will push (to the left or right)
        X = .Pos.X
        Y = .Pos.Y
        Select Case .Char.Heading

            Case NORTH
                For i = 1 To MoveTiles
                    HitTilesX(i) = X
                    HitTilesY(i) = Y - i
                Next i
                PushDir(1) = WEST
                PushDir(2) = EAST
                
            Case EAST
                For i = 1 To MoveTiles
                    HitTilesX(i) = X + i
                    HitTilesY(i) = Y
                Next i
                PushDir(1) = NORTH
                PushDir(2) = SOUTH
                
            Case SOUTH
                For i = 1 To MoveTiles
                    HitTilesX(i) = X
                    HitTilesY(i) = Y + i
                Next i
                PushDir(1) = WEST
                PushDir(2) = EAST
                
            Case WEST
                For i = 1 To MoveTiles
                    HitTilesX(i) = X - i
                    HitTilesY(i) = Y
                Next i
                PushDir(1) = NORTH
                PushDir(2) = SOUTH
                
            '  NW  N  NE
            '
            '  W   *   E
            '
            '  SW  S  SE
                
            Case NORTHEAST
                For i = 1 To MoveTiles
                    HitTilesX(i) = X + i
                    HitTilesY(i) = Y - i
                Next i
                PushDir(1) = NORTHWEST
                PushDir(2) = SOUTHEAST
                
            Case SOUTHEAST
                For i = 1 To MoveTiles
                    HitTilesX(i) = X + i
                    HitTilesY(i) = Y + i
                Next i
                PushDir(1) = SOUTHWEST
                PushDir(2) = NORTHEAST
                
            Case SOUTHWEST
                For i = 1 To MoveTiles
                    HitTilesX(i) = X - i
                    HitTilesY(i) = Y + i
                Next i
                PushDir(1) = NORTHWEST
                PushDir(2) = SOUTHEAST
                
            Case NORTHWEST
                For i = 1 To MoveTiles
                    HitTilesX(i) = X - i
                    HitTilesY(i) = Y - i
                Next i
                PushDir(1) = NORTHEAST
                PushDir(2) = SOUTHEAST
                
        End Select
        
        'Loop through the tiles the user will be hitting
        For i = 1 To 5
            
            'Make sure the tile is legal
            If HitTilesX(i) < 1 Then Exit For
            If HitTilesY(i) < 1 Then Exit For
            If HitTilesX(i) > MapInfo(.Pos.Map).Width Then Exit For
            If HitTilesY(i) > MapInfo(.Pos.Map).Height Then Exit For
            If MapInfo(.Pos.Map).Data(HitTilesX(i), HitTilesY(i)).Blocked > 0 Then Exit For
            If MapInfo(.Pos.Map).Data(HitTilesX(i), HitTilesY(i)).UserIndex > 0 Then Exit For
            
            'Check for a NPC on the tile
            NPCIndex = MapInfo(.Pos.Map).Data(HitTilesX(i), HitTilesY(i)).NPCIndex
            If NPCIndex > 0 Then
                
                'Only allow attackable NPCs to be pushed
                If NPCList(NPCIndex).Attackable Then
                    
                    'Push the NPC randomly in a push direction
                    NPC_Push NPCIndex, PushDir(Int(Rnd * 2) + 1), 4
                    
                End If
            
            End If
            
            'Check if the NPC moved away from the tile
            If MapInfo(.Pos.Map).Data(HitTilesX(i), HitTilesY(i)).NPCIndex = 0 Then
                
                'The user can go here, so move them
                .Flags.QuestNPC = 0
                .Flags.TradeWithNPC = 0
                .Flags.StepCounter = 0
                .Counters.MoveCounter = timeGetTime
                .Counters.StepsRan = 0
                MapInfo(.Pos.Map).Data(.Pos.X, .Pos.Y).UserIndex = 0
                .Pos.X = HitTilesX(i)
                .Pos.Y = HitTilesY(i)
                MapInfo(.Pos.Map).Data(.Pos.X, .Pos.Y).UserIndex = CasterIndex
                
            End If

        Next i
        
        'Check if the user moved
        If .Pos.X <> OldPos.X Or .Pos.Y <> OldPos.Y Then
            
            'Send the position update to the map and the skill packet
            'Position update
            ConBuf.PreAllocate 12
            ConBuf.Put_Byte DataCode.Server_WarpChar
            ConBuf.Put_Integer .Char.CharIndex
            ConBuf.Put_Byte .Pos.X
            ConBuf.Put_Byte .Pos.Y
            ConBuf.Put_Byte .Char.Heading
            'Skill packet
            ConBuf.Put_Byte DataCode.User_CastSkill
            ConBuf.Put_Byte SkID.Charge
            ConBuf.Put_Integer .Char.CharIndex
            ConBuf.Put_Byte OldPos.X
            ConBuf.Put_Byte OldPos.Y
            Data_Send ToMap, 0, ConBuf.Get_Buffer(), .Pos.Map
            
        End If
            
    End With

End Sub

Public Sub Skill_Warcry_PC(ByVal CasterIndex As Integer)

'*****************************************************************
'Reaver warcries, stunning all those in view
'*****************************************************************
Dim NPCIndex As Integer
Dim Map As Integer
Dim X As Integer
Dim Y As Integer

    'Store the map variable
    Map = UserList(CasterIndex).Pos.Map
    
    'Clear the conversion buffer (allocation is done later as NPCs to stun are found)
    ConBuf.Clear

    'Loop through the tiles near the Reaver
    For X = UserList(CasterIndex).Pos.X - MaxServerDistanceX To UserList(CasterIndex).Pos.X + MaxServerDistanceX
        For Y = UserList(CasterIndex).Pos.Y - MaxServerDistanceY To UserList(CasterIndex).Pos.Y + MaxServerDistanceY
            
            'Check for a valid tile
            If X > 0 Then
                If Y > 0 Then
                    If X <= MapInfo(Map).Width Then
                        If Y <= MapInfo(Map).Height Then

                            'Check for a valid NPC
                            NPCIndex = MapInfo(UserList(CasterIndex).Pos.Map).Data(X, Y).NPCIndex
                            If NPCIndex > 0 Then
                                If NPCList(NPCIndex).Flags.NPCAlive Then
                                    If NPCList(NPCIndex).Flags.NPCActive Then
                                        If NPCList(NPCIndex).Attackable Then
                                            If NPCList(NPCIndex).Hostile Then
                                                
                                                'Stun the NPC
                                                If NPCList(NPCIndex).ActiveSkills.StunTime - timeGetTime < 2000 Then
                                                    NPCList(NPCIndex).ActiveSkills.StunTime = timeGetTime + 2000
                                                    
                                                    'Create the status icon
                                                    ConBuf.Allocate 4
                                                    ConBuf.Put_Byte DataCode.Server_SetIcon
                                                    ConBuf.Put_Integer NPCList(NPCIndex).Char.CharIndex
                                                    ConBuf.Put_Byte IconID.Stun
                                                    
                                                End If
                                            
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        
                        End If
                    End If
                End If
            End If
        
        Next Y
    Next X
    
    If ConBuf.HasBuffer Then Data_Send ToMap, 0, ConBuf.Get_Buffer(), UserList(CasterIndex).Pos.Map, PP_StatusIcons

End Sub

Public Sub Skill_EF_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)
'*****************************************************************
'Reaver uses Exploding Finish, blowing up a character if the hit kills them
'*****************************************************************
Dim NPCIndex As Long
Dim Damage As Long
Dim X As Long
Dim Y As Long

    'Check for a valid target
    If NPCList(TargetIndex).Flags.NPCAlive = 0 Then Exit Sub
    If NPCList(TargetIndex).Flags.NPCActive = 0 Then Exit Sub
    If NPCList(TargetIndex).Attackable = 0 Then Exit Sub
    
    'Get the normal hit damage
    Damage = User_AttackNPC(CasterIndex, TargetIndex)
    
    With NPCList(TargetIndex)
        
        'Check if the hit will kill the NPC
        If .BaseStat(SID.MinHP) - Damage <= 0 Then
    
            'Explode!
            'Search for all near-by NPCs
            For X = .Pos.X - 1 To .Pos.X + 1
                For Y = .Pos.Y - 1 To .Pos.Y + 1
                    
                    'Check for a valid location
                    If X > 0 Then
                        If Y > 0 Then
                            If X <= MapInfo(.Pos.Map).Width Then
                                If Y <= MapInfo(.Pos.Map).Height Then
                                    
                                    'Check for a valid NPC
                                    NPCIndex = MapInfo(.Pos.Map).Data(X, Y).NPCIndex
                                    If NPCIndex > 0 Then
                                        With NPCList(NPCIndex)
                                            If .Flags.NPCAlive Then
                                                If .Flags.NPCActive Then
                                                    If .Attackable Then

                                                        'Apply splash damage
                                                        NPC_Damage NPCIndex, CasterIndex, .ModStat(SID.MaxHP) \ 4
                                                        
                                                    End If
                                                End If
                                            End If
                                        End With
                                    End If
                                    
                                End If
                            End If
                        End If
                    End If
                    
                Next Y
            Next X
    
        End If
        
    End With
    
    'Damage the primary NPC
    NPC_Damage TargetIndex, CasterIndex, Damage

End Sub

Public Sub Skill_Bash_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Reaver strikes an enemy hard, applying high damage and pushing them back
'*****************************************************************
Dim Damage As Long
Dim x1 As Integer
Dim x2 As Integer
Dim Y1 As Integer
Dim Y2 As Integer

    'Check for a valid target
    If NPCList(TargetIndex).Flags.NPCAlive = 0 Then Exit Sub
    If NPCList(TargetIndex).Flags.NPCActive = 0 Then Exit Sub
    If NPCList(TargetIndex).Attackable = 0 Then Exit Sub
    
    'Check for a valid distance
    x1 = UserList(CasterIndex).Pos.X
    Y1 = UserList(CasterIndex).Pos.Y
    x2 = NPCList(TargetIndex).Pos.X
    Y2 = NPCList(TargetIndex).Pos.Y
    If Abs(x1 - x2) > 1 Then Exit Sub
    If Abs(Y1 - Y2) > 1 Then Exit Sub
    
    'Attack the NPC
    Damage = User_AttackNPC(CasterIndex, TargetIndex) * 2.5
    NPC_Damage TargetIndex, CasterIndex, Damage
    
    'Push the NPC back
    NPC_Push TargetIndex, UserList(CasterIndex).Char.Heading, 5

End Sub

Public Sub Skill_Whirlwind_PC(ByVal CasterIndex As Integer)

'*****************************************************************
'Reaver pushes all enemies nearby away
'*****************************************************************
Dim NPCIndex As Integer
Dim Dir As Byte
Dim UserX As Long
Dim UserY As Long
Dim X As Long
Dim Y As Long

    'Get the user pos
    UserX = UserList(CasterIndex).Pos.X
    UserY = UserList(CasterIndex).Pos.Y
    
    'Loop through the 5x5 area around the user
    For X = -2 To 2
        For Y = -2 To 2
            
            'Check for a valid tile
            If UserX + X >= 1 Then
                If UserY + Y >= 1 Then
                    If UserX + X <= MapInfo(UserList(CasterIndex).Pos.Map).Width Then
                        If UserY + Y <= MapInfo(UserList(CasterIndex).Pos.Map).Height Then
                            
                            'Check for a NPC
                            NPCIndex = MapInfo(UserList(CasterIndex).Pos.Map).Data(UserX + X, UserY + Y).NPCIndex
                            If NPCIndex > 0 Then
                                
                                'Make sure the NPC is push-able
                                If NPCList(NPCIndex).Hostile Then
                                    If NPCList(NPCIndex).Attackable Then
                                        
                                        'Get the direction
                                        Dir = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(NPCIndex).Pos)
                                        
                                        'Push the NPC away
                                        NPC_Push NPCIndex, Dir, 5
                                        
                                    End If
                                End If
                            
                            End If
                        
                        End If
                    End If
                End If
            End If
        
        Next Y
    Next X

End Sub

Public Sub Skill_Rush_PCtoNPC(ByVal CasterIndex As Integer, ByVal TargetIndex As Integer)

'*****************************************************************
'Reaver rushes towards a target, inflicting great damage
'*****************************************************************
Dim Damage As Long
Dim NewPos As WorldPos
Dim TempHeading As Byte
Dim PosFound As Boolean
Dim CanUseHeading(1 To 8) As Byte
Dim BestHeadings(1 To 8) As Byte
Dim NewHeading As Byte
Dim X As Long
Dim Y As Long
Dim OldX As Byte
Dim OldY As Byte
Dim i As Long

    'Check for a valid target
    If NPCList(TargetIndex).Flags.NPCAlive = 0 Then Exit Sub
    If NPCList(TargetIndex).Flags.NPCActive = 0 Then Exit Sub
    If NPCList(TargetIndex).Attackable = 0 Then Exit Sub
    
    'If the user is right next to the NPC, perform a normal attack
    If Abs(CInt(UserList(CasterIndex).Pos.X) - CInt(NPCList(TargetIndex).Pos.X)) <= 1 Then
        If Abs(CInt(UserList(CasterIndex).Pos.Y) - CInt(NPCList(TargetIndex).Pos.Y)) <= 1 Then
            
            'Get the heading
            NewHeading = Server_FindDirection(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
            
            'Attack normally
            User_Attack CasterIndex, NewHeading
            
            Exit Sub
            
        End If
    End If
            
    'Make sure there is a free position near the NPC before anything
    'This will also find the positions we are allowed to use
    With NPCList(TargetIndex).Pos
        For X = .X - 1 To .X + 1
            For Y = .Y - 1 To .Y + 1
                If X > 0 Then
                    If Y > 0 Then
                        If X <= MapInfo(.Map).Width Then
                            If Y <= MapInfo(.Map).Height Then
                                If MapInfo(.Map).Data(X, Y).UserIndex = 0 Then
                                    If MapInfo(.Map).Data(X, Y).NPCIndex = 0 Then
                                        If MapInfo(.Map).Data(X, Y).Blocked = 0 Then
                                            PosFound = True
                                            
                                            Select Case Y
                                            
                                                'Above the NPC
                                                Case .Y - 1
                                                    If X = .X - 1 Then
                                                        CanUseHeading(NORTHWEST) = 1
                                                    ElseIf X = .X Then
                                                        CanUseHeading(NORTH) = 1
                                                    ElseIf X = .X + 1 Then
                                                        CanUseHeading(NORTHEAST) = 1
                                                    End If
                                                    
                                                'Same Y axis as NPC
                                                Case .Y
                                                    If X = .X - 1 Then
                                                        CanUseHeading(WEST) = 1
                                                    ElseIf X = .X + 1 Then
                                                        CanUseHeading(EAST) = 1
                                                    End If
                                                    
                                                'Below the NPC
                                                Case .Y + 1
                                                    If X = .X - 1 Then
                                                        CanUseHeading(SOUTHWEST) = 1
                                                    ElseIf X = .X Then
                                                        CanUseHeading(SOUTH) = 1
                                                    ElseIf X = .X + 1 Then
                                                        CanUseHeading(SOUTHEAST) = 1
                                                    End If
                                                    
                                            End Select
                                            
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next Y
        Next X
    End With
    If Not PosFound Then Exit Sub
        
    'Check for a valid path
    If Engine_ClearPath(UserList(CasterIndex).Pos.Map, UserList(CasterIndex).Pos.X, UserList(CasterIndex).Pos.Y, NPCList(TargetIndex).Pos.X, NPCList(TargetIndex).Pos.Y) = 0 Then Exit Sub

    'Find out the direction the user is heading towards the NPC
    TempHeading = Server_FindDirectionEX(UserList(CasterIndex).Pos, NPCList(TargetIndex).Pos)
    
    'Find the best position to take based off of the user's heading
    
    '  NW  N  NE
    '
    '  W   *   E
    '
    '  SW  S  SE
    
    BestHeadings(1) = TempHeading
    Select Case TempHeading
        Case NORTH
            BestHeadings(2) = NORTHEAST
            BestHeadings(3) = NORTHWEST
            BestHeadings(4) = EAST
            BestHeadings(5) = WEST
            BestHeadings(6) = SOUTHEAST
            BestHeadings(7) = SOUTHWEST
            BestHeadings(8) = SOUTH
        Case NORTHEAST
            BestHeadings(2) = NORTH
            BestHeadings(3) = EAST
            BestHeadings(4) = NORTHWEST
            BestHeadings(5) = SOUTHEAST
            BestHeadings(6) = WEST
            BestHeadings(7) = SOUTH
            BestHeadings(8) = SOUTHWEST
        Case EAST
            BestHeadings(2) = NORTHEAST
            BestHeadings(3) = SOUTHEAST
            BestHeadings(4) = NORTH
            BestHeadings(5) = SOUTH
            BestHeadings(6) = NORTHWEST
            BestHeadings(7) = SOUTHWEST
            BestHeadings(8) = WEST
        Case SOUTHEAST
            BestHeadings(2) = SOUTH
            BestHeadings(3) = EAST
            BestHeadings(4) = SOUTHWEST
            BestHeadings(5) = NORTHEAST
            BestHeadings(6) = WEST
            BestHeadings(7) = NORTH
            BestHeadings(8) = NORTHWEST
        Case SOUTH
            BestHeadings(2) = SOUTHWEST
            BestHeadings(3) = SOUTHEAST
            BestHeadings(4) = WEST
            BestHeadings(5) = EAST
            BestHeadings(6) = NORTHWEST
            BestHeadings(7) = NORTHEAST
            BestHeadings(8) = NORTH
        Case SOUTHWEST
            BestHeadings(2) = WEST
            BestHeadings(3) = SOUTH
            BestHeadings(4) = NORTHWEST
            BestHeadings(5) = SOUTHEAST
            BestHeadings(6) = NORTH
            BestHeadings(7) = EAST
            BestHeadings(8) = NORTHEAST
        Case WEST
            BestHeadings(2) = NORTHWEST
            BestHeadings(3) = SOUTHWEST
            BestHeadings(4) = NORTH
            BestHeadings(5) = SOUTH
            BestHeadings(6) = NORTHEAST
            BestHeadings(7) = SOUTHEAST
            BestHeadings(8) = EAST
        Case NORTHWEST
            BestHeadings(2) = NORTH
            BestHeadings(3) = WEST
            BestHeadings(4) = NORTHEAST
            BestHeadings(5) = SOUTHWEST
            BestHeadings(6) = EAST
            BestHeadings(7) = SOUTH
            BestHeadings(8) = SOUTHEAST
    End Select
            
    'Find the first best direction that can be used
    For i = 1 To 8
        If CanUseHeading(BestHeadings(i)) Then
            NewHeading = BestHeadings(i)
            Exit For
        End If
    Next i
    
    With UserList(CasterIndex)
    
        'Clear some values
        .Flags.QuestNPC = 0
        .Flags.TradeWithNPC = 0
        .Flags.StepCounter = 0
        .Counters.MoveCounter = timeGetTime
        .Counters.StepsRan = 0
        MapInfo(.Pos.Map).Data(.Pos.X, .Pos.Y).UserIndex = 0
        
        'Set the new position
        OldX = .Pos.X
        OldY = .Pos.Y
        NewPos = NPCList(TargetIndex).Pos
        Server_HeadToPos NewHeading, NewPos
        .Pos = NewPos
        MapInfo(.Pos.Map).Data(.Pos.X, .Pos.Y).UserIndex = CasterIndex
        
        'Set the user's new heading
        Select Case NewHeading
            Case NORTH
                NewHeading = SOUTH
            Case NORTHEAST
                NewHeading = SOUTHWEST
            Case EAST
                NewHeading = WEST
            Case SOUTHEAST
                NewHeading = NORTHWEST
            Case SOUTH
                NewHeading = NORTH
            Case SOUTHWEST
                NewHeading = NORTHEAST
            Case WEST
                NewHeading = EAST
            Case NORTHWEST
                NewHeading = SOUTHEAST
        End Select
        .Char.Heading = NewHeading
        
        'Send the position update to the map and the skill packet
        'Position update
        ConBuf.PreAllocate 12
        ConBuf.Put_Byte DataCode.Server_WarpChar
        ConBuf.Put_Integer .Char.CharIndex
        ConBuf.Put_Byte .Pos.X
        ConBuf.Put_Byte .Pos.Y
        ConBuf.Put_Byte .Char.Heading
        'Skill packet
        ConBuf.Put_Byte DataCode.User_CastSkill
        ConBuf.Put_Byte SkID.Rush
        ConBuf.Put_Integer .Char.CharIndex
        ConBuf.Put_Byte OldX
        ConBuf.Put_Byte OldY
        Data_Send ToMap, 0, ConBuf.Get_Buffer(), .Pos.Map
        
        'Calculate the damage
        Damage = User_AttackNPC(CasterIndex, TargetIndex)
        Damage = Damage * 2
        
        'Damage the NPC
        NPC_Damage TargetIndex, CasterIndex, Damage

    End With

End Sub

Public Function Skill_ValidSkillForClass(ByVal Class As Integer, ByVal SkillID As Byte) As Boolean

'*****************************************************************
'Check if the SkillID can be used by the class
'For skills with no defined requirements, theres no requirements
'Heal only has a requirement as an example
'*****************************************************************
Dim ClassReq As Integer

    'Sort by skill id
    Select Case SkillID
        Case SkID.Rush: ClassReq = Rush_ClassReq
        Case SkID.Whirlwind: ClassReq = Whirlwind_ClassReq
        Case SkID.Bash: ClassReq = Bash_ClassReq
        Case SkID.Warcry: ClassReq = Warcry_ClassReq
        Case SkID.Charge: ClassReq = Charge_ClassReq
        Case SkID.Grab: ClassReq = Grab_ClassReq
        Case SkID.ExplodingFinish: ClassReq = EF_ClassReq
        Case SkID.CrackArmor: ClassReq = CrackArmor_ClassReq
        Case SkID.Berserk: ClassReq = Berserk_ClassReq
        Case SkID.RageExplosion: ClassReq = RageExplosion_ClassReq
        Case SkID.Throw: ClassReq = Throw_ClassReq
        
        Case SkID.Hide: ClassReq = Hide_ClassReq
    End Select
    
    'Treat 0 as "all classes can use"
    If ClassReq <> 0 Then
    
        'Check the ClassReq VS the passed class
        Skill_ValidSkillForClass = (Class And ClassReq)

    Else

        'No requirements
        Skill_ValidSkillForClass = True

    End If

End Function
