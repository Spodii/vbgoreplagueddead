Attribute VB_Name = "Declares"
Option Explicit

'Used to record the number of packets coming in/out and what command ID they have
Public Const DEBUG_MapFPS As Boolean = True
Public Const DEBUG_UseLogging As Boolean = True                '//\\LOGLINE//\\
Public Const DEBUG_RecordPacketsOut As Boolean = True
Public Const DEBUG_RecordPacketsIn As Boolean = True
Public Const DEBUG_PrintPacketsIn As Boolean = False
Public Const DEBUG_PrintPacketsOut As Boolean = False

'********** Public CONSTANTS ***********

'Make the Objs.dat file
Public Const MakeObjsDat As Boolean = True

'Change to 1 to enable database optimization on runtime
Public Const OptimizeDatabase As Byte = 0

'If we run the server in high priority (recommended to use)
Public Const RunHighPriority As Byte = 0

'Screen resolution (must be identical to the values on the client!)
Public Const ScreenWidth As Long = 1024
Public Const ScreenHeight As Long = 768

'How long objects can be on the ground (in miliseconds) before being removed
Public Const GroundObjLife As Long = 300000 '5 minutes

'How long an object can remain in memory unused
Public Const ObjMemoryLife As Long = 300000 '5 minutes

'How long the maps last in memory when no users are on it
Public Const EmptyMapLife As Long = 180000  '3 minutes

'Amount of time that must elapse for certain user events (in miliseconds)
Public Const DelayTimeMail As Long = 3000   'Sending messages (has to be updated client-side, too)
Public Const DelayTimeTalk As Long = 500    'Talking (in any form)

'Change this value to add a cost to sending mail
Public Const MailCost As Long = 0

'Maximum allowed packets in per second from the client (used to prevent flooding)
Public Const MaxPacketsInPerSec As Long = 25    'During testing, I never got this over 12, so this should be a safe number

'Maximum amount of objects allowed on a single tile
Public Const MaxObjsPerTile As Byte = 5

'Running information
Public Const RunningSpeed As Byte = 20  'How much to increase speed when running
Public Const RunningCost As Long = 1    'How much stamina it cost to run

'Calculate the data in/out per sec or ont
Public Const CalcTraffic As Boolean = True

'Help (/help) and MOTD buffer
'These are filled in on frmMain.StartServer - it holds all the strings combined for faster sending
Public HelpBuffer() As Byte
Public MOTDBuffer() As Byte

'How many quests a user can accept at once
Public Const MaxQuests As Byte = 25

'ID of the sound effect used when no weapon is equipted and the user attacks
Public Const UnequiptedSwingSfx As Byte = 1

'Blocked directions - take the blocked byte and OR these values (If Blocked OR <Direction> Then...)
Public Const BlockedNorth As Byte = 1
Public Const BlockedEast As Byte = 2
Public Const BlockedSouth As Byte = 4
Public Const BlockedWest As Byte = 8
Public Const BlockedAll As Byte = 15

'Character types for CharList()
Public Const CharType_PC As Byte = 1
Public Const CharType_NPC As Byte = 2

'Client character types
Public Const ClientCharType_PC As Byte = 1
Public Const ClientCharType_NPC As Byte = 2
Public Const ClientCharType_Grouped As Byte = 3
Public Const ClientCharType_Slave As Byte = 4

'Max distance for two chars being on the same screen (for the rect distance) - values defined in tiles
Public Const MaxServerDistanceX As Long = ((ScreenWidth \ 32) \ 2) + 1
Public Const MaxServerDistanceY As Long = ((ScreenHeight \ 32) \ 2) + 1

'Group settings
Public Const Group_MaxUsers As Byte = 15        'The maximum number of users in a group
Public Const Group_MaxGroups As Byte = 255      'Max number of groups total
Public Const Group_InviteTime As Long = 20000   '(DO NOT raise this value above 255 seconds!) The time (in milliseconds) in which the user has to accept a group invitation
Public Const Group_MaxDistanceX As Byte = (ScreenWidth \ 32) * 2    'How far away you can be from the one who killed the NPC to get the exp
Public Const Group_MaxDistanceY As Byte = (ScreenHeight \ 32) * 2   'Values are defined in tiles

'************ Positioning ************
Type WorldPos   'Holds placement information
    Map As Integer  'Map
    X As Byte       'X coordinate
    Y As Byte       'Y coordinate
End Type

'************ Object types ************
Public Const MaxObjAmount As Integer = 100          'Maximum number you can hold of an item (if Stacking <= 0)
Public Const MAX_INVENTORY_SLOTS As Byte = 49       'Maximum number of inventory slots
Public Type udtObjData
    Name As String                  'Name
    ObjType As Byte                 'Type (armor, weapon, item, etc)
    ClassReq As Integer             'Class requirement
    GrhIndex As Long                'Graphic index
    SpriteBody As Integer           'Index of the body sprite to change to
    SpriteWeapon As Integer         'Index of the weapon sprite to change to
    SpriteHair As Integer           'Index of the hair sprite to change to
    SpriteHead As Integer           'Index of the head sprite to change to
    SpriteWings As Integer          'Index of the wings sprite to change to
    WeaponType As Byte              'What type of weapon, if it is a weapon
    WeaponRange As Byte             'Range of the weapon (only applicable if a ranged WeaponType)
    UseGrh As Long                  'Grh of the object when used (projectile for ranged, slash for melee, effects for use-once)
    UseSfx As Byte                  'Sound effect played when the object is used (based on the .wav's number)
    ProjectileRotateSpeed As Byte   'How fast the projectile rotates (if at all)
    Value As Long                   'Value of the object
    RepHP As Long                   'How much HP to replenish
    RepEP As Long                   'How much EP to replenish
    RepHPP As Integer               'Percentage of HP to replenish
    RepEPP As Integer               'Percentage of EP to replenish
    ReqStr As Long                  'Required strength to use the item
    ReqAgi As Long                  'Required agility
    ReqInt As Long                  'Required intelligence
    Stacking As Integer             'How much the item can be stacked up (-1 for no limit, 0 for
    AddStat(FirstModStat To NumStats) As Long   'How much to add to the stat by the SID
    Pointer As Integer
End Type
Public ObjData As ObjData
Public Type Obj 'Holds info about a object
    ObjIndex As Integer     'Index of the object
    Amount As Integer       'Amount of the object
End Type

'************ Trading ************
Public Const TRADESTATE_CLOSED As Byte = 0      'Trading is closed - table not used
Public Const TRADESTATE_TRADING As Byte = 1     'Objects are still being placed, no confirm yet
Public Const TRADESTATE_ACCEPT As Byte = 2      'The user has accepted what they have placed, no items can be added / removed
Public Const TRADESTATE_FINISHED As Byte = 3    'The user has confirmed the trade - can only happen after accepting
Public Type TradeTableObj
    UserInvSlot As Byte
    Amount As Integer
End Type
Public Type TradeTable
    User1 As Integer
    User2 As Integer
    Objs1(1 To 9) As TradeTableObj
    Objs2(1 To 9) As TradeTableObj
    Gold1 As Long
    Gold2 As Long
    User1State As Byte
    User2State As Byte
End Type
Public TradeTable() As TradeTable
Public NumTradeTables As Byte

'************ Map Tiles/Information ************
Type NPCLoadData    'Used to load NPCs from the .temp map files
    NPCNum As Integer
    X As Byte
    Y As Byte
End Type
Type MapBlock   'Information for each map block
    '*** IMPORTANT! *** ADDING ARRAYS TO THIS UDT WILL BREAK THE LOADING!
    'If you must add an array, try adding it to a different UDT, like I did with the objects
    TileExitMap As Integer      'Warp location when user touches the tile
    TileExitX As Byte
    TileExitY As Byte
    Blocked As Byte             'If the tile is blocked
    Mailbox As Byte             'If there is a mailbox on the tile
    UserIndex As Integer        'Index of the user on the tile
    NPCIndex As Integer         'Index of the NPC on the tile
    LightIntensity As Byte      'Intensity of the lighting on a scale of 0 to 10
End Type
Type ObjBlock   'Information on an object on a map block
    NumObjs As Byte             'Number of objects on the tile
    ObjLife() As Long           'When the object was created (used to determine it's life)
    ObjInfo() As Obj            'Information of the object on the tile
End Type
Type MapInfo    'Map information
    NumUsers As Integer     'Number of users on the map
    Name As String          'Name of the map
    MapVersion As Integer   'Version of the map
    Width As Byte           'Dimensions of the map
    Height As Byte
    Weather As Byte         'What weather effects the map has going
    Job As Byte
    PVP As Byte
    Music As Byte           'The music file number of the map
    DataLoaded As Byte      'If the map data is loaded
    UnloadTimer As Long     'How long until the surface unloads
    Data() As MapBlock      'Holds the information on each tile; Data(TileX, TileY)
    ObjTile() As ObjBlock   'Holds the information on the objects on the tiles; Obj(TileX, TileY)
End Type
Public MapInfo() As MapInfo

'************ Mailing System ************
Public Const MaxMailPerUser As Byte = 50    'How much mail each user may have maximum
Public Const MaxMailObjs As Byte = 10       'How many objects can be attached to a message maximum
Type MailData   'Mailing system information
    Subject As String
    WriterName As String
    RecieveDate As Date
    Message As String
    New As Byte
    Obj(1 To MaxMailObjs) As Obj
End Type

'************ Group/Party System ************
Type GroupData
    Users() As Integer
    NumUsers As Byte
End Type
Public GroupData() As GroupData
Public NumGroups As Integer         'Holds the number of groups created (empty groups included, so this is the highest group index)
Public NumEmptyGroups As Integer    'Holds the number of empty groups - used for group index recycling to speed up the process

'************ Generic Character Data ************
Type CharData   'Charlist types (for reverting from CharIndex to PC/NPC index)
    Index As Integer
    CharType As Byte    '0 = Unused, 1 = PC, 2 = NPC
End Type
Public CharList() As CharData

'************ Quest ************
Public Type Quest
    Name As String                  'The quest's name
    StartTxt As String              'What the NPC says to the player to explain the crisis
    AcceptTxt As String             'What the NPC says when the player accepts the quest
    IncompleteTxt As String         'What the NPC says to the player when they return without completing the quest
    FinishTxt As String             'What the NPC says when the player finishes the quest
    AcceptReqObj As Integer         'Index of the object the user is required to have to accept
    AcceptReqObjAmount As Integer   'How much of the object the user must have before accepting
    AcceptReqFinishQuest As Integer 'Quest that must be finished prior to accepting this quest
    AcceptRewExp As Long            'Amount of Exp the user gets for accepting the quest
    AcceptRewGold As Long           'Amount of gold the user gets for accepting the quest
    AcceptRewObj As Integer         'Object the user gets for accepting the quest
    AcceptRewObjAmount As Integer   'Amount of the object the user gets for accepting the quest
    AcceptLearnSkill As Byte        'Skill the user learns for accepting the quest (by SkID value)
    FinishReqObj As Integer         'Object the user must have to finish the quest
    FinishReqObjAmount As Integer   'Amount of the object the user must have to finish the quest
    FinishReqNPC As Integer         'Index of the NPC the user must kill to finish the quest
    FinishReqNPCAmount As Integer   'How many of the NPCs the user must kill to finish the quest
    FinishRewExp As Long            'Exp the user gets for finishing the quest
    FinishRewGold As Long           'How much gold the user gets for finishing the quest
    FinishRewObj As Integer         'The index of the object the user gets for finishing the quest
    FinishRewObjAmount As Integer   'The amount of the object the user gets for finishing the quest
    FinishLearnSkill As Byte        'Skill the user learns for finishing the quest (by SkID value)
    Redoable As Byte                'If the quest can be done infinite times
End Type
Public QuestData() As Quest

'************ NPC/Character types ************
Type Char   'Holds data for a user or NPC character
    CharIndex As Integer        'Character's index
    Hair As Integer             'Hair index
    Head As Integer             'Head index
    Body As Integer             'Body index
    Weapon As Integer           'Weapon index
    Wings As Integer            'Wings index
    Heading As Byte             'Current direction facing
    HeadHeading As Byte         'Direction char's head is facing
End Type
Public Type QuestStatus 'Status of user's current quests
    NPCKills As Integer     'How many of the targeted NPCs the user has killed
End Type
Type UserFlags  'Flags for a user
    UserLogged As Byte      'If the user is logged in
    LastViewedMail As Byte  'The last mail index which the user viewed
    TradeWithNPC As Integer 'NPC the user is trading with
    TargetIndex As Integer  'Index of the NPC or Player targeted
    Target As Byte          'Type of targeting - 0 for none, 1 for player, 2 for NPC
    GMLevel As Byte         'What type of admin the user is: 0 = None, 1 = GM
    Disconnecting As Byte   'If the user will be disconnected after data is sent
    QuestNPC As Integer     'The ID of the NPC that the user is talking to about a quest
    InviteGroup As Byte     'The index of the group the user has been invited to
    DoNotSave As Byte       'Used in special cases to define to bypass saving
    TradeTable As Byte      'The trade table the user is in (0 for none)
    TradeWith As Integer    'The user index the user wants to trade with
    CreatedStats As Byte    'If the user's stats object was created
    StepCounter As Byte     'How many steps the user has taken since MoveCounter value was set
    GrabChar As Integer     'The index of the char the user is grabbing, if any
    Berserk As Byte         'If the user is in "berserk" mode (1 or 0)
    LastAttacked As Long    'The time the user was last attacked at
    Hiding As Byte          'If the user has Hide being used
End Type
Type UserCounters   'Counters for a user
    IdleCount As Long           'Stores last time the user sent an action packet
    LastPacket As Long          'Stores last time the user sent ANY packet
    AttackCounter As Long       'Delay time for attacks
    SkillCounter As Long        'Delay time for skills
    MoveCounter As Long         'Stores last time the user moved
    SpellExhaustion As Long     'Time until another spell can be casted
    DelayTimeMail As Long       'Mail write delay time
    DelayTimeTalk As Long       'Talk delay time
    PacketsInCount As Long      'Packets in per second (used to prevent packet flooding)
    PacketsInTime As Long       'When the packet counting started
    GroupCounter As Long        'How long the user has to accept to join a group
    SwapCounter As Long         'Time the user must wait to use the /swap command again
    StepsRan As Byte            'How many steps the user has ran for the StepCounter
    RageCounter As Long         'Rage counter
End Type
Type UserOBJ    'Objects the user has
    ObjIndex As Long    'Index of the object
    Amount As Long      'Amount of the objects
    Equipped As Byte    'If the object is equipted
End Type
Type ActiveSkills 'Skills casted on a user / NPC
    StunTime As Long
    CrackArmorTime As Long
End Type
Type Cache_Server_MakeChar
    Body As Integer
    Head As Integer
    Heading As Byte
    X As Byte
    Y As Byte
    Speed As Byte
    Name As String
    Weapon As Integer
    Hair As Integer
    Wings As Integer
    HP As Byte
    EP As Byte
    ChatID As Byte
    CharType As Byte
End Type
Type PacketCache
    Server_MakeChar As Cache_Server_MakeChar
End Type
Type User   'Holds data for a user
    Account As String       'Name of the user's account
    Name As String          'Name of the user
    Char As Char            'Defines users looks
    Desc As String          'User's description
    Pos As WorldPos         'User's current position
    Class As Integer        'User's class
    BankGold As Long        'Amount of gold the user has in the bank
    LoginTime As Long       'The time the user logged in (in server ticks)
    BytesIn As Long         'Total bandwidth in from server
    BytesOut As Long        'Total bandwidth out from server
    GroupIndex As Integer   'The index of the group the user is part of (if any)
    SendBuffer() As Byte    'Buffer for sending data
    BufferSize As Long      'Size of the buffer
    HasBuffer As Byte       'If there is anything in the buffer
    LastPacketSent As Long  'Time the last packet was sent
    PPCount As Long         'Packet priority count-down (only valid if PPValue = PP_Low)
    PacketWait As Long      'Packet wait count-down (not to be confused with the packet priority - this one is for Packet_WaitTime)
    Object(1 To MAX_INVENTORY_SLOTS) As UserOBJ 'The user's inventory
    Bank(1 To MAX_INVENTORY_SLOTS) As Obj       'The user's bank items
    inChannel(1 To NumChannels) As Byte         'If the user is part of the channel of the same ID (1 = yes, 0 = no)
    WeaponEqpObjIndex As Integer    'The index of the equipted weapon
    WeaponEqpSlot As Byte           'Slot of the equipted weapon
    WeaponType As Byte              'Type of weapon the user is using
    ArmorEqpObjIndex As Integer     'The index of the equipted armor
    ArmorEqpSlot As Byte            'Slot of the equipted armorn
    WingsEqpObjIndex As Integer     'The index of the equipted Wings
    WingsEqpSlot As Byte            'Slot of the equipted Wings
    Counters As UserCounters        'Declares the user counters
    Stats As UserStats              'Declares the user stats
    Flags As UserFlags              'Declares the user Flags
    ActiveSkills As ActiveSkills    'Declares the skills casted on the user
    PacketCache As PacketCache      'Typical users do NOT need to worry about the packet cache at all and just let it do its job
    NumSlaves As Byte               'Number of "slave" (ie summoned or pet) NPCs the user has
    SlaveNPCIndex() As Integer      'NPC index of the slaves (not CharIndex since you can't slave a PC)
    KnownSkills(1 To NumSkills) As Byte         'Declares the skills known by the user
    NumCompletedQuests As Integer               'The total number of quests that were completed by the user (Ubound of CompletedQuests)
    CompletedQuests() As Integer                'Each index of the byte contains the ID of a quest completed
    Quest(1 To MaxQuests) As Integer            'The quest index of the current quests if any
    QuestStatus(1 To MaxQuests) As QuestStatus  'Counts certain parts of quests that require being counted (ie NPC kills)
    MailID(1 To MaxMailPerUser) As Long         'ID of the user's mail
    MailboxPos As WorldPos                      'Position of the last-used mailbox
End Type
Public UserList() As User   'Holds data for each user
Type NPCFlags   'Flags for a NPC
    NPCAlive As Byte        'If the NPC is alive and visible
    NPCActive As Byte       'If the NPC is active
    Thralled As Byte        'If the NPC is thralled (if so, it does not get saved or respawn)
    UpdateStats As Byte     'If to update the mod stats
    GrabbedBy As Integer    'Index of the character grabbing the NPC
    RendedBy As Integer     'Index of the character who rended the NPC
    IsIdling As Byte        'If the NPC is just sitting around doing nothing (their AI is not making them move/attack/etc)
    WasIdling As Byte       'Holds if in the last frame, the NPC was idling, in opposed to IsIdling, which holds the current frame
End Type
Type NPCCounters    'Counters for a NPC
    RespawnCounter As Long      'Stores the death time to respawn later (if a summoned/thralled NPC, its how long until they die off)
    SpellExhaustion As Long     'Time until another spell can be casted
    ActionDelay As Long         'How long until the NPC can perform another action
    ReplenishCounter As Long    'How long it has been since the NPC has last attacked or been attacked
    LastConcussion As Long      'When the NPC got their last concussion
    LastRend As Long            'When the NPC got their last rend
    RendDamage As Long          'Damage caused each time from rend
    NumRends As Byte            'Number of rends left to perform on the NPC
    EmoDelay As Long            'Delay until the NPC can use another emoticon
End Type
Type NPC    'Holds all the NPC variables
    Name As String  'Name of the NPC
    Char As Char    'Defines NPC looks
    Desc As String  'Description
    Pos As WorldPos         'Current NPC Postion
    StartPos As WorldPos    'Spawning location of the NPC
    NPCNumber As Integer    'The NPC index within NPC.dat
    AI As Byte              'Used AI algorithm
    ChatID As Byte          'Chat ID used
    RespawnWait As Long     'How long for the NPC to respawn
    Attackable As Byte      'If the NPC is attackable
    Hostile As Byte         'If the NPC is hostile
    GiveGLD As Long         'How much gold given on death
    Quest As Integer        'Quest index
    AttackRange As Byte     'The NPC's attack range
    AttackGrh As Long       'Grh used when the NPC attacks
    AttackSfx As Byte       'Sound effect played when the NPC attacks
    OwnerIndex As Integer   'The user index the NPC is bound to (ie summoned or a pet)
    ProjectileRotateSpeed As Byte   'If a projectile, how fast it rotates
    ActiveSkills As ActiveSkills    'Declares the skills casted on the NPC
    Flags As NPCFlags               'Declares the NPC's Flags
    Counters As NPCCounters         'Declares the NPC's counters
    NumVendItems As Byte            'Number of items the NPC is vending
    NumDropItems As Byte            'Number of items the NPC is dropping
    BaseStat(1 To NumStats) As Long 'Declares the NPC's stats
    ModStat(FirstModStat To NumStats) As Long   'Declares the NPC's stats
    
    'THESE ARRAYS MUST STAY DOWN HERE AT THE BOTTOM OF THE UDT!
    VendItems() As Obj              'Information on the item the NPC is vending
    DropItems() As Obj              'Information on the item to drop
    DropRate() As Single            'The drop rate of the item in the DropItems() array sharing the same index
End Type
Public NPCList() As NPC     'Holds data for each NPC

'Two bytes put together (used for the NPC loading/saving of vending/drop item amounts)
Public Type NPCBytes
    Vend As Byte
    Drop As Byte
End Type

'Server information
Public Type ServerInfo
    IIP As String           'Internal IP of the server (used to bind to the correct IP)
    EIP As String           'External IP of the server (used to identify the server from another server)
    ServerPort As String    'Port used to communicate between servers (server <-> server)
    Port As Integer         'Port used to communicate to clients (server <-> client)
End Type
Public ServerInfo() As ServerInfo
Public LocalSocketID As Long    'Index of the socket ID of this server
Public MaxUsers As Integer      'Max users allowed on this server
Public ServerMap() As Byte      'Points to the server that is uses the map
Public NumServers As Byte       'Total number of servers
Public ServerID As Byte         'The ID of this server (ServerID in Server.ini file)

'***********************************
'********** Misc Values ************
'***********************************
'All the below can be changed without worry of conversion

'Weapon type constants
Public Enum WeaponType
    Hand = 0        'Weapon is hand-based
    Staff = 1       'Weapon is a staff
    Dagger = 2      'Weapon is a dagger
    Sword = 3       'Weapon is a sword
    Throwing = 4    'Weapon is thrown (ninja stars, throwing knives, etc)
End Enum
#If False Then
Private Hand, Staff, Dagger, Sword, Throwing
#End If

'Object types
Public Const OBJTYPE_USEONCE As Byte = 1    'Objects that can be used only once
Public Const OBJTYPE_WEAPON As Byte = 2     'Weapons of all types
Public Const OBJTYPE_ARMOR As Byte = 3      'Body armors
Public Const OBJTYPE_WINGS As Byte = 4      'Wings
Public Const OBJTYPE_USEINFINITE As Byte = 5    'USEONCE that does not vanish after usage

'Constants for headings
Public Const NORTH As Byte = 1
Public Const EAST As Byte = 2
Public Const SOUTH As Byte = 3
Public Const WEST As Byte = 4
Public Const NORTHEAST As Byte = 5
Public Const SOUTHEAST As Byte = 6
Public Const SOUTHWEST As Byte = 7
Public Const NORTHWEST As Byte = 8

'********** Public VARS ***********

Public ResPos As WorldPos
Public StartPos As WorldPos
Public NumUsers As Integer  'Current number of users
Public LastUser As Integer  'Index of the last user
Public LastChar As Integer
Public LastNPC As Integer
Public NumNPCs As Integer
Public NumMaps As Integer
Public NumQuests As Integer
Public NumObjDatas As Integer

Public IdleLimit As Long
Public LastPacket As Long

'All the users located on a map
Public Type MapUsersType
    Index() As Long
End Type
Public MapUsers() As MapUsersType

'Names of NPCs (for NPCs involved in quests)
Public NPCName() As String

'Number of connections (used just for displaying purposes)
Public CurrConnections As Long

'The time the server started (in system time)
Public ServerStartTime As Long

'States the server is running
Public ServerRunning As Byte

'Buffer used for conversions to send to Data_Send
Public ConBuf As DataBuffer

'Traffic information (bytes are converted to kbytes to allow larger numbers)
Public DataIn As Long
Public DataOut As Long
Public DataKBIn As Long
Public DataKBOut As Long

'Server FPS tracking (DEBUG_MapFPS)
Public Type ServerFPS
    FPS As Long         'FPS
    Users As Integer    'Number of users
    NPCs As Integer     'Number of NPCs
End Type
Public ServerFPSUbound As Long
Public ServerFPS() As ServerFPS
Public FPSIndex As Long

'The number of bytes we need to send all of our known skills
Public NumBytesForSkills As Long

'Server is unloading
Public UnloadServer As Byte

'Lets us know if User_CleanArray needs to be called - can NOT be called in the middle of a loop!
Public CallUserCleanArray As Byte

'Maximum number of NPCs allowed at once per server
'This value should not be raised without raising the datatype of NPC indexes from integer to long (not recommended)
'If this value is too low, try decreasing the map unloading time so NPCs are unloaded quicker
Public Const MaxNPCs As Integer = 32765 'Note this is a little less then 2 ^ 16, just in case

'Packet messages that are cached so they don't have to built real-time
Public Type CachedMessage
    Data() As Byte
End Type
Public cMessage() As CachedMessage

Public DebugPacketsOut() As Long
Public DebugPacketsIn() As Long

'Keep alive packet
Public KeepAlivePacket() As Byte

'Replenish timer
Public Const NPCReplenishTime As Long = 15000

'Flag if the NPC stats need to be raised
Public RecoverNPCStats As Boolean

Public Const StatMax As Long = 1000
Public StatCost(0 To StatMax) As Long
Public Const LevelMax As Long = 500
Public LevelCost(0 To LevelMax) As Long

'********** EXTERNAL FUNCTIONS ***********
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Public Declare Function timeGetTimeEX Lib "winmm.dll" Alias "timeGetTime" () As Long
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
