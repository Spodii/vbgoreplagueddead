Attribute VB_Name = "Groups"
Option Explicit

Public Sub Group_AddUser(ByVal UserIndex As Integer, ByVal GroupIndex As Byte)

'*****************************************************************
'Adds a user to an existing group
'*****************************************************************
Dim i As Long

    'Check for valid group information
    If GroupIndex > NumGroups Then Exit Sub
    If GroupIndex < 1 Then Exit Sub
    If GroupData(GroupIndex).NumUsers = 0 Then Exit Sub
    If GroupData(GroupIndex).NumUsers >= Group_MaxUsers Then
        ConBuf.PreAllocate 3 + Len(UserList(GroupData(GroupIndex).Users(1)).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 106
        ConBuf.Put_String UserList(GroupData(GroupIndex).Users(1)).Name
        Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
        Exit Sub
    End If
    
    'Add the user to the group
    GroupData(GroupIndex).NumUsers = GroupData(GroupIndex).NumUsers + 1
    ReDim Preserve GroupData(GroupIndex).Users(1 To GroupData(GroupIndex).NumUsers)
    GroupData(GroupIndex).Users(GroupData(GroupIndex).NumUsers) = UserIndex
    UserList(UserIndex).GroupIndex = GroupIndex
    
    'Join group message and tell the user that just joined who else is in the group
    ConBuf.PreAllocate 3 + Len(UserList(GroupData(GroupIndex).Users(1)).Name) + ((GroupData(GroupIndex).NumUsers - 1) * 4)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 107
    ConBuf.Put_String UserList(GroupData(GroupIndex).Users(1)).Name
    If GroupData(GroupIndex).NumUsers > 1 Then
        For i = 1 To GroupData(GroupIndex).NumUsers - 1
            ConBuf.Put_Byte DataCode.Server_ChangeCharType
            ConBuf.Put_Integer UserList(GroupData(GroupIndex).Users(i)).Char.CharIndex
            ConBuf.Put_Byte ClientCharType_Grouped
        Next i
    End If
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
    'Tell the group members the user joined and change the char type
    ConBuf.PreAllocate 7 + Len(UserList(UserIndex).Name)
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 108
    ConBuf.Put_String UserList(UserIndex).Name
    ConBuf.Put_Byte DataCode.Server_ChangeCharType
    ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
    ConBuf.Put_Byte ClientCharType_Grouped
    Data_Send ToGroup, UserIndex, ConBuf.Get_Buffer
    
End Sub

Public Sub Group_SplitEXP(ByVal UserIndex As Integer, ByVal GroupIndex As Integer, ByVal EXP As Long)

'*****************************************************************
'Splits up the EXP among the group members in range
'*****************************************************************
Dim GiveUsers() As Integer
Dim NumUsersInRange As Byte
Dim tIndex As Integer
Dim i As Byte

    'Check for a valid group
    If GroupIndex = 0 Then Exit Sub
    
    'Check for a valid amount of exp
    If EXP = 0 Then Exit Sub
    
    'Default to give to all users
    ReDim GiveUsers(1 To GroupData(GroupIndex).NumUsers)
    
    'Loop through all the users
    For i = 1 To GroupData(GroupIndex).NumUsers
    
        'Hold the index in a smaller variable
        tIndex = GroupData(GroupIndex).Users(i)
    
        'Confirm that it is a valid index
        If tIndex > 0 Then
        
            'Check if the user is on the same map
            If UserList(UserIndex).Pos.Map = UserList(tIndex).Pos.Map Then
            
                'Check the distance
                If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, Group_MaxDistanceX, Group_MaxDistanceY) Then
                    
                    'The user is in range, set them in the array to get the exp
                    NumUsersInRange = NumUsersInRange + 1
                    GiveUsers(NumUsersInRange) = tIndex
                    
                End If
                
            End If
        
        End If
        
    Next i
    
    'Split up the EXP
    EXP = (EXP * (1 + ((NumUsersInRange - 1) \ 5))) \ NumUsersInRange
    If EXP < 1 Then EXP = 1

    'Give the exp to all the users
    For i = 1 To NumUsersInRange
        User_RaiseExp GiveUsers(i), EXP
    Next i
    
    'Clear the GiveUsers array
    Erase GiveUsers()

End Sub

Public Sub Group_SplitGold(ByVal UserIndex As Integer, ByVal GroupIndex As Integer, ByVal Gold As Long)

'*****************************************************************
'Splits up the Gold among the group members in range
'*****************************************************************
Dim GiveUsers() As Integer
Dim NumUsersInRange As Byte
Dim tIndex As Integer
Dim i As Byte

    'Check for a valid group
    If GroupIndex = 0 Then Exit Sub
    
    'Check for a valid amount of gold
    If Gold = 0 Then Exit Sub
    
    'Default to give to all users
    ReDim GiveUsers(1 To GroupData(GroupIndex).NumUsers)
    
    'Loop through all the users
    For i = 1 To GroupData(GroupIndex).NumUsers
    
        'Hold the index in a smaller variable
        tIndex = GroupData(GroupIndex).Users(i)
    
        'Confirm that it is a valid index
        If tIndex > 0 Then
        
            'Check if the user is on the same map
            If UserList(UserIndex).Pos.Map = UserList(tIndex).Pos.Map Then
            
                'Check the distance
                If Server_RectDistance(UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y, UserList(tIndex).Pos.X, UserList(tIndex).Pos.Y, Group_MaxDistanceX, Group_MaxDistanceY) Then
                    
                    'The user is in range, set them in the array to get the exp
                    NumUsersInRange = NumUsersInRange + 1
                    GiveUsers(NumUsersInRange) = tIndex
                    
                End If
                
            End If
        
        End If
        
    Next i
    
    'Split up the Gold
    Gold = (Gold \ NumUsersInRange) + 1

    'Give the gold to all the users
    For i = 1 To NumUsersInRange
        UserList(GiveUsers(i)).Stats.BaseStat(SID.Gold) = UserList(GiveUsers(i)).Stats.BaseStat(SID.Gold) + Gold
    Next i
    
    'Clear the GiveUsers array
    Erase GiveUsers()

End Sub

Public Sub Group_RemoveUser(ByVal UserIndex As Integer, ByVal GroupIndex As Integer)

'*****************************************************************
'Removes a user from an existing group
'*****************************************************************
Dim i As Byte
Dim j As Byte

    'Check for valid group information
    If GroupIndex > NumGroups Then Exit Sub
    If GroupIndex < 1 Then Exit Sub
    If GroupData(GroupIndex).NumUsers = 0 Then Exit Sub 'Group deleted

    'Tell the user they have left the group and change all current group members to not group members for the UserIndex
    ConBuf.PreAllocate (4 * GroupData(GroupIndex).NumUsers) + 2
    ConBuf.Put_Byte DataCode.Server_Message
    ConBuf.Put_Byte 109
    For i = 1 To GroupData(GroupIndex).NumUsers
        ConBuf.Put_Byte DataCode.Server_ChangeCharType
        ConBuf.Put_Integer UserList(GroupData(GroupIndex).Users(i)).Char.CharIndex
        ConBuf.Put_Byte ClientCharType_PC
    Next i
    Data_Send ToIndex, UserIndex, ConBuf.Get_Buffer
    
    'Check how many people are left in the group
    If GroupData(GroupIndex).NumUsers > 1 Then
    
        'Tell everyone else they have left
        ConBuf.PreAllocate 7 + Len(UserList(UserIndex).Name)
        ConBuf.Put_Byte DataCode.Server_Message
        ConBuf.Put_Byte 110
        ConBuf.Put_String UserList(UserIndex).Name
        ConBuf.Put_Byte DataCode.Server_ChangeCharType
        ConBuf.Put_Integer UserList(UserIndex).Char.CharIndex
        ConBuf.Put_Byte ClientCharType_PC
        Data_Send ToGroup, UserIndex, ConBuf.Get_Buffer
        
        'Clear the user's group flag
        UserList(UserIndex).GroupIndex = 0
    
    Else
    
        'This is the last person, so just empty the group out
        GroupData(GroupIndex).NumUsers = 0
        Erase GroupData(GroupIndex).Users
        
        'Raise the empty group count
        NumEmptyGroups = NumEmptyGroups + 1
        
        'Clear the user's group flag
        UserList(UserIndex).GroupIndex = 0
        
        Exit Sub
        
    End If
    
    'Find the slot the user has in the group
    For i = 1 To GroupData(GroupIndex).NumUsers
    
        'Index found
        If GroupData(GroupIndex).Users(i) = UserIndex Then

            'If the user is the last one in the group, just resize the array
            If GroupData(GroupIndex).NumUsers = i Then
                GroupData(GroupIndex).NumUsers = GroupData(GroupIndex).NumUsers - 1
                ReDim Preserve GroupData(GroupIndex).Users(1 To GroupData(GroupIndex).NumUsers)
                Exit Sub
            End If
            
            'The user is not at the end of the array, and theres more then one user
            For j = i To GroupData(GroupIndex).NumUsers - 1
                
                'Move all the users in the group down in the array to fill the now empty slow
                GroupData(GroupIndex).Users(j) = GroupData(GroupIndex).Users(j + 1)
                
            Next j
            
            'Remove the left-over slot at the end
            GroupData(GroupIndex).NumUsers = GroupData(GroupIndex).NumUsers - 1
            ReDim Preserve GroupData(GroupIndex).Users(1 To GroupData(GroupIndex).NumUsers)
            Exit Sub
            
        End If
        
    Next i
    
End Sub

Public Function Group_Create(ByVal UserIndex As Integer) As Byte

'*****************************************************************
'Find the next free group index and return it
'*****************************************************************
Dim i As Byte

    'Check if theres any groups yet
    If NumGroups = 0 Then
        NumGroups = 1
        Group_Create = 1
        ReDim GroupData(1 To 1) As GroupData
        Exit Function
    End If
    
    'Check if there are any empty groups - if so, find out which index it is
    If NumEmptyGroups > 0 Then
        For i = 1 To NumGroups
            If GroupData(i).NumUsers = 0 Then
                
                'Found a group not in use, use it
                NumEmptyGroups = NumEmptyGroups - 1 'We took one of the empty groups
                Group_Create = i
                Exit Function
                
            End If
        Next i
    End If
    
    'No groups found, check if we can make a new one
    If NumGroups + 1 >= Group_MaxGroups Then
        Data_Send ToIndex, UserIndex, cMessage(111).Data()
        Exit Function
    End If
    
    'Add the new group
    NumGroups = NumGroups + 1
    ReDim Preserve GroupData(1 To NumGroups)
    Group_Create = NumGroups

End Function

