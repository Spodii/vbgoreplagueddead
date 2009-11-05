Attribute VB_Name = "DataIDs"
Option Explicit

'********** Status Icons ************
Public Const NumIcons As Byte = 3
Public Type IconID
    CrackArmor As Byte
    Stun As Byte
    Berserk As Byte
End Type
Public IconID As IconID

'********** Emoticons ************
Public Const NumChannels As Byte = 4
Public Type ChannelID
    Gossip As Byte
    Help As Byte
    GroupSearch As Byte
    Market As Byte
End Type
Public ChannelID As ChannelID

'********** Emoticons ************
Public Const NumEmotes As Byte = 11
Public Type EmoID
    Dots As Byte
    Exclimation As Byte
    Question As Byte
    Surprised As Byte
    Heart As Byte
    Hearts As Byte
    HeartBroken As Byte
    Utensils As Byte
    Meat As Byte
    ExcliQuestion As Byte
    Sweat As Byte
End Type
Public EmoID As EmoID

'********** Classes ************
'Classes work by using bitwise operations, so each class ID must be a power of 2 (1, 2, 4, 8, 16, 32, 64, or 128)
'If you want more clases, change the classes to "Integer"
'To set class requirements, OR the values together
'EX:
'ClassReq = Warrior OR Rogue
'This means the class must be a Warrior or Rogue
'To check the values, use AND
'EX:
'If ClassReq AND UserClass Then
'   User meets requirements
'Else
'   User doesn't meet requirements
'End If
Public Type ClassID
    Civilian As Integer
    Reaver As Integer
    Infiltrator As Integer
    Engineer As Integer
    SquadLeader As Integer
    Job As Integer
    NoReq As Integer
End Type
Public ClassID As ClassID

'********** Packets ************
'Data String Codenames (Reduces all data transfers to 1 byte tags)
Public Type DataCode
    Comm_Talk As Byte
    Comm_TalkChannel As Byte
    Comm_Emote As Byte
    Comm_Whisper As Byte
    Comm_GroupTalk As Byte
    Comm_FontType_Talk As Byte
    Comm_FontType_Fight As Byte
    Comm_FontType_Info As Byte
    Comm_FontType_Quest As Byte
    Comm_FontType_Group As Byte
    Comm_UseBubble As Byte  'Do not use this alone - OR it onto Comm_Talk!
    Server_MailMessage As Byte
    Server_MailBox As Byte
    Server_MailItemTake As Byte
    Server_MailItemRemove As Byte
    Server_MailDelete As Byte
    Server_MailCompose As Byte
    Server_UserCharIndex As Byte
    Server_SetUserPosition As Byte
    Server_MakeChar As Byte
    Server_MakeCharCached As Byte
    Server_EraseChar As Byte
    Server_MoveChar As Byte
    Server_ChangeChar As Byte
    Server_MakeObject As Byte
    Server_EraseObject As Byte
    Server_PlaySound As Byte
    Server_PlaySound3D As Byte
    Server_Who As Byte
    Server_CharHP As Byte
    Server_CharEP As Byte
    Server_SetIcon As Byte
    Server_RemoveIcon As Byte
    Server_SetCharDamage As Byte
    Server_Help As Byte
    Server_Disconnect As Byte
    Server_Connect As Byte
    Server_Message As Byte
    Server_SetCharSpeed As Byte
    Server_MakeProjectile As Byte
    Server_MakeSlash As Byte
    Server_MailObjUpdate As Byte
    Server_MakeEffect As Byte
    Server_SendQuestInfo As Byte
    Server_ChangeCharType As Byte
    Server_KeepAlive As Byte
    Server_WarpChar As Byte
    Map_LoadMap As Byte
    Map_DoneLoadingMap As Byte
    Map_SendName As Byte
    User_Target As Byte
    User_JoinChannel As Byte
    User_LeaveChannel As Byte
    User_KnownSkills As Byte
    User_Attack As Byte
    User_SetInventorySlot As Byte
    User_Desc As Byte
    User_Login As Byte
    User_Get As Byte
    User_Drop As Byte
    User_Use As Byte
    User_Move As Byte
    User_Rotate As Byte
    User_LeftClick As Byte
    User_RightClick As Byte
    User_LookLeft As Byte
    User_LookRight As Byte
    User_Blink As Byte
    User_ChannelList As Byte
    User_SetRage As Byte
    User_Trade_StartNPCTrade As Byte
    User_Trade_BuyFromNPC As Byte
    User_Trade_SellToNPC As Byte
    User_Trade_Trade As Byte
    User_Trade_UpdateTrade As Byte
    User_Trade_Accept As Byte
    User_Trade_Finish As Byte
    User_Trade_RemoveItem As Byte
    User_Trade_Cancel As Byte
    User_Hide As Byte
    User_Bank_Open As Byte
    User_Bank_PutItem As Byte
    User_Bank_TakeItem As Byte
    User_Bank_UpdateSlot As Byte
    User_Bank_Balance As Byte
    User_Bank_Deposit As Byte
    User_Bank_Withdraw As Byte
    User_BaseStat As Byte
    User_ModStat As Byte
    User_CastSkill As Byte
    User_ChangeInvSlot As Byte
    User_Emote As Byte
    User_StartQuest As Byte
    User_CancelQuest As Byte
    User_SetWeaponRange As Byte
    User_RequestMakeChar As Byte
    User_RequestUserCharIndex As Byte
    User_ChangeServer As Byte
    User_ConfirmPosition As Byte
    User_Group_Make As Byte
    User_Group_Join As Byte
    User_Group_Leave As Byte
    User_Group_Invite As Byte
    User_Group_Info As Byte
    User_Profile As Byte
    User_ChangeClass As Byte
    User_SendClass As Byte
    User_SetSkillDelay As Byte
    GM_Approach As Byte
    GM_Summon As Byte
    GM_Kick As Byte
    GM_Raise As Byte
    GM_SetGMLevel As Byte
    GM_Thrall As Byte
    GM_DeThrall As Byte
    GM_Warp As Byte
    GM_FindItem As Byte
    GM_SQL As Byte
    GM_GiveSkill As Byte
    GM_GiveGold As Byte
    GM_GiveObject As Byte
    GM_KillMap As Byte
    GM_Kill As Byte
    GM_WarpToMap As Byte
    GM_IPInfo As Byte
    Combo_ProjectileSoundRotateDamage As Byte
    Combo_SoundRotateDamage As Byte
    Combo_SlashSoundRotateDamage As Byte
End Type
Public DataCode As DataCode

'********** Character Stats/Skills ************
Public Type StatOrder
    'These can NOT be modded (theres no ModStat())
    MinHP As Byte
    MinEP As Byte
    Gold As Byte
    Points As Byte
    EXP As Byte
    ELU As Byte
    ELV As Byte
    'These CAN be modded (ModStat() is used)
    MaxHIT As Byte
    MinHIT As Byte
    DEF As Byte
    MaxHP As Byte
    MaxEP As Byte
    Str As Byte
    Agi As Byte     'For NPCs, this is the hit rate
    Inte As Byte
    Speed As Byte
    Dex As Byte
    Brave As Byte
    WeaponSkill As Byte
    Armor As Byte
    Accuracy As Byte
    Evade As Byte
    Perception As Byte
    Regen As Byte
    Recov As Byte
    Tactics As Byte
    Immunity As Byte
    Rage As Byte
    Concussion As Byte
    Rend As Byte
    Bloodlust As Byte
    AttackDelay As Byte
    Stealth As Byte
    CriticalAttack As Byte
    SpeedInfil As Byte
    Thievery As Byte
End Type
Public SID As StatOrder 'Stat ID
Public Const NumStats As Byte = 36
Public Const FirstModStat As Byte = 8   'The lowest number of the first stat that can be modded

Public Type SkillID
    Rush As Byte
    Bash As Byte
    Whirlwind As Byte
    Warcry As Byte
    Charge As Byte
    Grab As Byte
    CrackArmor As Byte
    ExplodingFinish As Byte
    Throw As Byte
    Berserk As Byte
    RageExplosion As Byte
    
    Hide As Byte
    Cloak As Byte
    MeatCar As Byte
    Stun As Byte
    Blink As Byte
    Grapple As Byte
    BackStab As Byte
    MirrorImage As Byte
    Flash As Byte
    Mark As Byte
    Slow As Byte
    EagleEye As Byte
    Trap As Byte
    ChargedAttack As Byte
End Type
Public SkID As SkillID  'Skill IDs
Public Const NumSkills As Byte = 25

Public Sub InitDataCommands()

    'Load the values for the data commands
    With IconID
        .CrackArmor = 1
        .Stun = 2
        .Berserk = 3
    End With
    
    With ChannelID
        .Gossip = 1
        .GroupSearch = 2
        .Help = 3
        .Market = 4
    End With
    
    With EmoID
        .Dots = 1
        .Exclimation = 2
        .Question = 3
        .Surprised = 4
        .Heart = 5
        .Hearts = 6
        .HeartBroken = 7
        .Utensils = 8
        .Meat = 9
        .ExcliQuestion = 10
        .Sweat = 11
    End With

    With SkID
        .Bash = 1
        .Rush = 2
        .Berserk = 3
        .Charge = 4
        .CrackArmor = 5
        .ExplodingFinish = 6
        .Grab = 7
        .RageExplosion = 8
        .Throw = 9
        .Warcry = 10
        .Whirlwind = 11
        
        .Hide = 12
        .Cloak = 13
        .MeatCar = 14
        .Stun = 15
        .Blink = 16
        .Grapple = 17
        .BackStab = 18
        .MirrorImage = 19
        .Flash = 20
        .Mark = 21
        .Slow = 22
        .EagleEye = 23
        .Trap = 24
        .ChargedAttack = 25
    End With
    
    With ClassID
        'These values must be based off of powers of 2! (Note: The 16th bit is not 2 ^ 16, its -(2 ^ 15) because its signed)
        .Civilian = 1       '2 ^ 0
        .Reaver = 2         '2 ^ 1
        .Engineer = 4       '2 ^ 2 ... etc
        .Infiltrator = 8
        .SquadLeader = 16
        .Job = 32
        
        'This sets every bit to 1, which means that it will work with every class
        .NoReq = -1 'Read up on how signed binary works if you want to figure out why this is -1
        
    End With

    With SID
        'These can NOT be modded (theres no ModStat())
        .MinHP = 1
        .MinEP = 2
        .Gold = 3
        .Points = 4
        .EXP = 5
        .ELU = 6
        .ELV = 7
        'These CAN be modded (ModStat() is used)
        .MaxHIT = 8
        .MinHIT = 9
        .Immunity = 10
        .DEF = 11
        .MaxHP = 12
        .MaxEP = 13
        .Tactics = 14
        .Str = 15
        .Agi = 16
        .Inte = 17
        .Speed = 18
        .Dex = 19
        .Brave = 20
        .WeaponSkill = 21
        .Armor = 22
        .Accuracy = 23
        .Evade = 24
        .Perception = 25
        .Regen = 26
        .Recov = 27
        .Rage = 28
        .Concussion = 29
        .Rend = 30
        .Bloodlust = 31
        .AttackDelay = 32
        .Stealth = 33
        .CriticalAttack = 34
        .SpeedInfil = 35
        .Thievery = 36
    End With

    With DataCode
        .User_RequestMakeChar = 1
        .GM_Thrall = 2
        .User_SetRage = 3
        .Comm_TalkChannel = 4
        .Server_UserCharIndex = 5
        .Comm_Emote = 6
        .Server_SetUserPosition = 7
        .Map_LoadMap = 8
        .Map_DoneLoadingMap = 9
        .GM_Raise = 10
        .GM_Kick = 11
        .Server_CharHP = 12
        .GM_Summon = 13
        .User_ChangeServer = 14
        .Map_SendName = 15
        .User_Attack = 16
        .Server_MakeChar = 17
        .Server_EraseChar = 18
        .Server_MoveChar = 19
        .Server_ChangeChar = 20
        .Server_MakeObject = 21
        .Server_EraseObject = 22
        .User_KnownSkills = 23
        .User_SetInventorySlot = 24
        .User_StartQuest = 25
        .Server_Connect = 26
        .Server_PlaySound = 27
        .User_Login = 28
        .User_ChannelList = 29
        .Comm_Whisper = 30
        .Server_Who = 31
        .User_Move = 32
        .User_Rotate = 33
        .User_LeftClick = 34
        .User_RightClick = 35
        .User_Group_Info = 36
        .User_Get = 37
        .User_Drop = 38
        .User_Use = 39
        .GM_Approach = 40
        .Comm_Talk = 41
        .Server_SetCharDamage = 42
        .User_ChangeInvSlot = 43
        .User_Emote = 44
        .Server_CharEP = 45
        .Server_Disconnect = 46
        .User_LookLeft = 47
        .User_LookRight = 48
        .User_Blink = 49
        .User_Trade_RemoveItem = 50
        .User_Trade_BuyFromNPC = 51
        .User_BaseStat = 52
        .User_ModStat = 53
        .User_Hide = 54
        ' = 55
        .Server_SendQuestInfo = 56
        .User_ConfirmPosition = 57
        .Server_Help = 58
        .User_Desc = 59
        .User_Trade_Cancel = 60
        .User_Target = 61
        .User_Trade_StartNPCTrade = 62
        .User_Trade_SellToNPC = 63
        .User_CastSkill = 64
        .Server_WarpChar = 65
        .Server_SetIcon = 66
        .Server_RemoveIcon = 67
        .User_SetSkillDelay = 68
        .User_JoinChannel = 69
        .User_LeaveChannel = 70
        .Server_MailBox = 71
        .Server_MailMessage = 72
        .User_RequestUserCharIndex = 73
        .Server_MailItemTake = 74
        .Server_MailObjUpdate = 75
        .Server_MailDelete = 76
        .Server_MailCompose = 77
        .GM_SetGMLevel = 78
        .Server_Message = 79
        .GM_DeThrall = 80
        .Server_PlaySound3D = 81
        .Server_SetCharSpeed = 82
        .User_SetWeaponRange = 83
        .Server_MakeProjectile = 84
        .Server_MakeSlash = 85
        .Server_MakeEffect = 86
        .User_Bank_Open = 87
        .User_Bank_PutItem = 88
        .User_Bank_TakeItem = 89
        .User_Bank_UpdateSlot = 90
        .User_Group_Join = 91
        .User_Group_Invite = 92
        .User_Group_Leave = 93
        .User_Group_Make = 94
        .Comm_GroupTalk = 95
        .User_Bank_Deposit = 96
        .User_Bank_Withdraw = 97
        .User_Bank_Balance = 98
        .GM_Warp = 99
        .Server_ChangeCharType = 100
        .User_Trade_Trade = 101
        .User_Trade_UpdateTrade = 102
        .User_Trade_Accept = 104
        .User_Trade_Finish = 105
        .User_CancelQuest = 106
        .Combo_ProjectileSoundRotateDamage = 107
        .Combo_SoundRotateDamage = 108
        .Combo_SlashSoundRotateDamage = 109
        .User_ChangeClass = 110
        .Server_MakeCharCached = 111
        .GM_FindItem = 112
        .GM_SQL = 113
        .GM_GiveSkill = 114
        .GM_GiveObject = 115
        .GM_KillMap = 116
        .GM_Kill = 117
        .GM_WarpToMap = 118
        .GM_IPInfo = 119
        ' = 120
        .GM_GiveGold = 121
        .User_Profile = 122
        .User_SendClass = 123
        .Server_KeepAlive = 124
                
        'This values can be used over again since they aren't used in their own packet header
        .Comm_FontType_Fight = 1
        .Comm_FontType_Info = 2
        .Comm_FontType_Quest = 3
        .Comm_FontType_Talk = 4
        .Comm_FontType_Group = 5

        'Value 128 can be used over again since this does not count as an ID in itself - just ignore this variable! ;)
        .Comm_UseBubble = 128
        
    End With

End Sub

