Attribute VB_Name = "Declares"
Option Explicit

'Stat costs
Public Const StatMax As Long = 1000
Public StatCost(0 To StatMax) As Long

'Prevents us getting effects (namely blood splatters) before the map even loads, which will set them in the wrong place
Public AcceptEffects As Boolean

'********** Debug/Display Settings ************
'These are your key constants - reccomended you turn off ALL debug constants before
' compiling your code for public usage just speed reasons

'Set this to true to force updater check
Public Const ForceUpdateCheck As Boolean = False

'Running speed - make sure you have the same value on the server!
Public Const RunningSpeed As Byte = 20
Public Const RunningCost As Long = 1    'How much stamina it cost to run

'Max chat bubble width
Public Const BubbleMaxWidth As Long = 140

'********** Objects **********
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
    ReqLvl As Long                  'Required level
    Stacking As Integer             'How much the item can be stacked up (-1 for no limit, 0 for
    AddStat(FirstModStat To NumStats) As Long   'How much to add to the stat by the SID
    Pointer As Integer
End Type
Public ObjData() As udtObjData
Public NumObjDatas As Integer

'********** NPC chat info ************
Public Type NPCChatLineCondition
    Condition As Byte           'The condition used (see NPCCHAT_COND_)
    Value As Long               'Used to hold a numeric condition value
    ValueStr As String          'Used to hold a value for SAY conditions
End Type
Public Type NPCChatLine
    NumConditions As Byte       'Total number of conditions
    Conditions() As NPCChatLineCondition
    Text As String              'The text that will be said
    Style As Byte               'The style used for the text (see NPCCHAT_STYLE_)
    Delay As Integer            'The delay time applied after saying this line
End Type
Public Type NPCChatAskAnswer    'The individual chat input answers
    Text As String              'The answer string
    GotoID As Byte              'ID the answer will move to
End Type
Public Type NPCChatAskLine      'Individual chat input lines
    Question As String          'The question text
    NumAnswers As Byte          'Number of answers that can be used
    Answer() As NPCChatAskAnswer
    AskFlags() As Integer
    NumAskFlags As Integer
End Type
Public Type NPCChatAsk          'Chat input information (ASK parameters)
    StartAsk As Byte            'ID to start the asking on
    Ask() As NPCChatAskLine     'Holds all the ASK questions/responses
End Type
Public Type NPCChat
    Format As Byte              'Format of the chat (see NPCCHAT_FORMAT_)
    ChatLine() As NPCChatLine   'The information on the chat line
    NumLines As Byte            'The number of chat lines
    Distance As Long            'The distance the user must be from the NPC to activate the chat
    Ask As NPCChatAsk           'All the ASK information
End Type
Public NPCChat() As NPCChat

'Conditions (this are used as bit-flags, so only use powers of 2!)
Public Const NPCCHAT_COND_LEVELLESSTHAN As Long = 2 ^ 0
Public Const NPCCHAT_COND_LEVELMORETHAN As Long = 2 ^ 1
Public Const NPCCHAT_COND_HPLESSTHAN As Long = 2 ^ 2
Public Const NPCCHAT_COND_HPMORETHAN As Long = 2 ^ 3
Public Const NPCCHAT_COND_KNOWSKILL As Long = 2 ^ 4
Public Const NPCCHAT_COND_DONTKNOWSKILL As Long = 2 ^ 5
Public Const NPCCHAT_COND_SAY As Long = 2 ^ 6

'Chat formats
Public Const NPCCHAT_FORMAT_RANDOM As Byte = 0
Public Const NPCCHAT_FORMAT_LINEAR As Byte = 1

'Chat sytles
Public Const NPCCHAT_STYLE_BOTH As Byte = 0
Public Const NPCCHAT_STYLE_BOX As Byte = 1
Public Const NPCCHAT_STYLE_BUBBLE As Byte = 2

'Client character types
Public Const ClientCharType_PC As Byte = 1
Public Const ClientCharType_NPC As Byte = 2
Public Const ClientCharType_Grouped As Byte = 3
Public Const ClientCharType_Slave As Byte = 4

'********** Trade table ************
Public Type TradeObj
    Amount As Long
    ObjIndex As Integer
End Type
Public Type TradeTable
    User1Name As String              'The name of the table
    User2Name As String
    User1Accepted As Byte
    User2Accepted As Byte
    Trade1(1 To 9) As TradeObj  'The objects both indexes have entered
    Trade2(1 To 9) As TradeObj
    Gold1 As Long               'The gold both indexes have entered
    Gold2 As Long
    MyIndex As Byte             'States whether this client is index 1 or 2
End Type
Public TradeTable As TradeTable

'********** Other stuff ************
Public BaseStats(1 To NumStats) As Long
Public ModStats(FirstModStat To NumStats) As Long
Public UserClass As Integer
Public UserRage As Long

'Delay timers for packet-related actions (so to not spam the server)
Public Const AttackDelay As Long = 200  'These constants are client-side only
Public Const LootDelay As Long = 500    ' - changing these lower wont make it faster server-side!
Public LastAttackTime As Long
Public LastLootTime As Long

'Cached packets
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
    MP As Byte
    ChatID As Byte
    CharType As Byte
End Type
Type PacketCache
    Server_MakeChar As Cache_Server_MakeChar
End Type
Public PacketCache As PacketCache

'Item description variables
Public ItemDescWidth As Long
Public ItemDescLine() As String
Public ItemDescLines As Byte

'Object constants
Public Const MAX_INVENTORY_SLOTS As Byte = 49

'Active ASK information
Public Type ActiveAsk
    AskName As String
    AskIndex As Byte
    ChatIndex As Byte
    QuestionTxt As String
End Type
Public ActiveAsk As ActiveAsk

'User's inventory
Type Inventory
    ObjIndex As Long
    Amount As Integer
    Equipped As Boolean
End Type

'Quest information
Type QuestInfo
    Name As String
    Desc As String
End Type
Public QuestInfo() As QuestInfo
Public QuestInfoUBound As Byte

'Messages
Public NumMessages As Byte
Public Message() As String

'Signs
Public Signs() As String

'Known user skills/spells
Public UserKnowSkill(1 To NumSkills) As Byte

'Attack range
Public UserAttackRange As Byte

'User status vars
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory
Public UserBank(1 To MAX_INVENTORY_SLOTS) As Inventory

'The time the last packet from the server arrived
Public LastServerPacketTime As Long

'If there is a clear path to the target (if any)
Public ClearPathToTarget As Byte

'Skill delay time
Public SkillDelayTimeStart As Long
Public SkillDelayTimeEnd As Long

Public sndBuf As DataBuffer
Public ChatBufferChunk As Single
Public SoxID As Long
Public GettingAccount As Boolean
Public SocketMoveToIP As String
Public SocketMoveToPort As Integer
Public SocketOpen As Byte
Public TargetCharIndex As Integer
Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

'Mail sending spam prevention
Public LastMailSendTime As Long

'Holds the skin the user is using at the time
Public CurrentSkin As String

'Blocked directions - take the blocked byte and OR these values (If Blocked OR <Direction> Then...)
Public Const BlockedNorth As Byte = 1
Public Const BlockedEast As Byte = 2
Public Const BlockedSouth As Byte = 4
Public Const BlockedWest As Byte = 8
Public Const BlockedAll As Byte = 15

Public UseSfx As Byte
Public UseMusic As Byte

'States if the project is unloading (has to give Sox time to unload)
Public IsUnloading As Byte

'User login information
Public UserPassword As String
Public UserName As String
Public UserBody As Byte
Public UserHead As Byte

'Holds the name of the last person to whisper to the client
Public LastWhisperName As String

'Zoom level - 0 = No Zoom, > 0 = Zoomed
Public ZoomLevel As Single
Public Const MaxZoomLevel As Single = 0.3

'Cursor flash rate
Public Const CursorFlashRate As Long = 450

'If click-warping is on or not (can only be used by GMs)
Public UseClickWarp As Byte

'Emoticon delay
Public EmoticonDelay As Long

'How long char remains aggressive-faced after being attacked
Public Const AGGRESSIVEFACETIME = 4000

'Save password check
Public SavePass As Boolean

'Maximum variable sizes
Public Const MAXLONG As Long = (2 ^ 31) - 1
Public Const MAXINT As Integer = (2 ^ 15) - 1
Public Const MAXBYTE As Byte = (2 ^ 8) - 1

'********** DLL CALLS ***********
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
