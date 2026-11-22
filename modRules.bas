Attribute VB_Name = "modRules"
'==============================================================
' Grid Gobbler - rules, word categories and square generation
'==============================================================
Option Explicit

' Rule kinds
Public Const RK_MULTIPLES As Long = 0
Public Const RK_FACTORS As Long = 1
Public Const RK_PRIMES As Long = 2
Public Const RK_EQUALS As Long = 3
Public Const RK_CATEGORY As Long = 4
Public Const RK_EVENODD As Long = 5
Public Const RK_COMPARE As Long = 6
Public Const RK_SQUARES As Long = 7

' Game modes
Public Const MODE_MATH As Long = 1
Public Const MODE_WORDS As Long = 2
Public Const MODE_MIXED As Long = 3

Public Type GameRule
    Kind As Long        ' RK_xxx
    N As Long           ' main parameter
    MaxV As Long        ' largest number used on the board
    Ops As Long         ' equals: 1=+ 2=+- 3=+-x 4=+-x/   compare: 0=less 1=greater
    Cat As Long         ' category index for word rules
    Title As String     ' text shown at the top of the screen
End Type

Public Type GridCell
    Txt As String       ' what is printed in the square
    Value As Long       ' numeric value (math rules)
    Good As Boolean     ' True = fits the rule
    Filled As Boolean   ' False = already gobbled / empty
End Type

Private Const MAX_CATS As Long = 80

Private mItems() As String
Private mItemCount As Long
Private mNumCats As Long
Private mBuiltin As Long
Private mCatTitle(0 To MAX_CATS - 1) As String
Private mCatOne(0 To MAX_CATS - 1) As String
Private mCatClue(0 To MAX_CATS - 1) As String
Private mCatStart(0 To MAX_CATS - 1) As Long
Private mCatLen(0 To MAX_CATS - 1) As Long
Private mTrickStart(0 To MAX_CATS - 1) As Long
Private mTrickLen(0 To MAX_CATS - 1) As Long
Private mCatUsed(0 To MAX_CATS - 1) As Boolean
Private mLastKind As Long

'--------------------------------------------------------------
' Built-in word lists. The optional last list holds near-misses
' (Spider is not an insect...). More categories can be added
' without recompiling - see packs\sample-pack.txt
'--------------------------------------------------------------
Public Sub InitRules()
    mItemCount = 0
    mNumCats = 0
    ReDim mItems(0 To 0)

    AddCat "Fruits", "a fruit", _
        "Apple,Banana,Cherry,Grape,Mango,Lemon,Pear,Melon,Papaya,Coconut,Guava,Pineapple,Strawberry,Blueberry", ""
    AddCat "Vegetables", "a vegetable", _
        "Carrot,Potato,Onion,Celery,Pea,Spinach,Radish,Turnip,Leek,Beet,Lettuce,Cabbage,Broccoli,Garlic,Kale", ""
    AddCat "Colors", "a color", _
        "Red,Blue,Green,Yellow,Purple,Pink,Brown,Black,White,Gray,Violet,Indigo,Maroon,Beige,Crimson,Scarlet", ""
    AddCat "Planets", "a planet", _
        "Mercury,Venus,Earth,Mars,Jupiter,Saturn,Uranus,Neptune", "Moon,Sun,Pluto,Comet"
    AddCat "Mammals", "a mammal", _
        "Dog,Cat,Horse,Whale,Bat,Tiger,Lion,Mouse,Rabbit,Bear,Goat,Camel,Zebra,Otter,Panda,Dolphin,Monkey", _
        "Shark,Lizard,Frog,Snake,Turtle"
    AddCat "Birds", "a bird", _
        "Eagle,Robin,Owl,Parrot,Penguin,Crow,Duck,Swan,Hawk,Pigeon,Ostrich,Falcon,Goose,Sparrow,Heron,Toucan", ""
    AddCat "Insects", "an insect", _
        "Ant,Bee,Beetle,Wasp,Moth,Butterfly,Ladybug,Termite,Flea,Hornet,Mosquito,Firefly", _
        "Spider,Scorpion,Centipede,Worm,Snail"
    AddCat "Musical Instruments", "a musical instrument", _
        "Piano,Guitar,Drum,Violin,Flute,Trumpet,Harp,Cello,Tuba,Banjo,Oboe,Clarinet,Trombone,Ukulele", ""
    AddCat "Sports", "a sport", _
        "Soccer,Tennis,Golf,Hockey,Rugby,Boxing,Baseball,Skiing,Rowing,Judo,Bowling,Archery,Karate,Surfing,Volleyball", ""
    AddCat "Countries", "a country", _
        "France,Japan,Brazil,Canada,Egypt,India,Mexico,Peru,Kenya,Spain,Italy,China,Norway,Ireland,Sweden,Germany", _
        "Paris,Tokyo,London,Africa,Europe,Asia"
    AddCat "Shapes", "a shape", _
        "Circle,Square,Triangle,Oval,Hexagon,Star,Cube,Sphere,Cone,Pyramid,Rectangle,Octagon,Cylinder,Pentagon", ""
    AddCat "Weather", "a kind of weather", _
        "Rain,Snow,Hail,Fog,Sleet,Wind,Storm,Thunder,Drizzle,Frost,Tornado,Blizzard,Hurricane,Lightning", ""
    AddCat "Body Parts", "a body part", _
        "Elbow,Knee,Ankle,Wrist,Nose,Ear,Thumb,Heart,Lung,Liver,Skull,Spine,Shoulder,Finger,Tongue", ""
    AddCat "Vehicles", "a vehicle", _
        "Car,Bus,Truck,Train,Bicycle,Tractor,Taxi,Van,Scooter,Tram,Airplane,Helicopter,Submarine,Rocket,Sailboat", ""
    AddCat "Clothing", "an item of clothing", _
        "Shirt,Jacket,Sock,Hat,Scarf,Glove,Sweater,Skirt,Dress,Boot,Coat,Jeans,Vest,Shorts,Pajamas", ""
    AddCat "Trees", "a tree", _
        "Oak,Maple,Pine,Birch,Willow,Cedar,Elm,Spruce,Redwood,Aspen,Hickory,Cypress,Sequoia,Beech,Poplar", ""
    AddCat "Flowers", "a flower", _
        "Tulip,Daisy,Lily,Orchid,Sunflower,Daffodil,Poppy,Marigold,Petunia,Carnation,Pansy,Peony,Lotus,Begonia,Zinnia", ""
    AddCat "Tools", "a tool", _
        "Hammer,Wrench,Saw,Drill,Pliers,Chisel,Shovel,Rake,Screwdriver,Axe,Trowel,Crowbar,Sander,Clamp", ""
    AddCat "Furniture", "a piece of furniture", _
        "Chair,Table,Sofa,Desk,Bed,Dresser,Bookcase,Stool,Bench,Couch,Cabinet,Wardrobe,Crib,Armchair", ""
    AddCat "Fish", "a fish", _
        "Trout,Tuna,Cod,Carp,Pike,Herring,Sardine,Catfish,Flounder,Halibut,Minnow,Guppy,Swordfish,Perch,Eel", _
        "Squid,Jellyfish,Starfish,Lobster"
    AddCat "Months", "a month", _
        "January,February,March,April,May,June,July,August,September,October,November,December", ""
    AddCat "Days of the Week", "a day of the week", _
        "Monday,Tuesday,Wednesday,Thursday,Friday,Saturday,Sunday", "Today,Weekend,Holiday"
    AddCat "Kitchen Things", "something from the kitchen", _
        "Spoon,Fork,Plate,Bowl,Cup,Kettle,Pan,Spatula,Whisk,Ladle,Toaster,Oven,Blender,Colander,Grater", ""
    AddCat "Jobs", "a job", _
        "Doctor,Teacher,Farmer,Pilot,Chef,Nurse,Plumber,Dentist,Lawyer,Baker,Painter,Firefighter,Carpenter,Sailor,Barber", ""
    AddCat "U.S. States", "a U.S. state", _
        "Texas,Ohio,Maine,Utah,Iowa,Idaho,Florida,Alaska,Hawaii,Oregon,Nevada,Kansas,Vermont,Montana,Arizona", _
        "Chicago,Boston,Dallas,Denver"
    AddCat "Languages", "a language", _
        "English,Spanish,French,German,Italian,Chinese,Japanese,Arabic,Russian,Hindi,Greek,Latin,Swedish,Korean,Dutch", ""
    AddCat "Drinks", "a drink", _
        "Water,Milk,Juice,Tea,Coffee,Lemonade,Cocoa,Soda,Smoothie,Cider,Milkshake", ""
    AddCat "Desserts", "a dessert", _
        "Cake,Pie,Cookie,Brownie,Pudding,Donut,Cupcake,Muffin,Fudge,Sundae,Custard,Waffle", ""

    mBuiltin = mNumCats
    ResetRuleHistory
End Sub

Private Sub AddCat(ByVal catTitle As String, ByVal oneOf As String, ByVal members As String, ByVal tricks As String, _
                   Optional ByVal clue As String = "")
    If mNumCats >= MAX_CATS Then Exit Sub
    mCatTitle(mNumCats) = catTitle
    mCatOne(mNumCats) = oneOf
    mCatClue(mNumCats) = clue
    mCatStart(mNumCats) = mItemCount
    mCatLen(mNumCats) = AddItems(members)
    mTrickStart(mNumCats) = mItemCount
    mTrickLen(mNumCats) = AddItems(tricks)
    mNumCats = mNumCats + 1
End Sub

Private Function AddItems(ByVal itemList As String) As Long
    Dim parts() As String, i As Long, w As String, n As Long
    If Len(itemList) = 0 Then Exit Function
    parts = Split(itemList, ",")
    For i = 0 To UBound(parts)
        w = Trim$(parts(i))
        If Len(w) > 0 Then
            ReDim Preserve mItems(0 To mItemCount)
            mItems(mItemCount) = w
            mItemCount = mItemCount + 1
            n = n + 1
        End If
    Next
    AddItems = n
End Function

'--------------------------------------------------------------
' Category packs: every .txt file in the packs folder.
'   [Dinosaurs]
'   is=a dinosaur
'   clue=They ruled the Earth millions of years ago.
'   yes=Triceratops,Stegosaurus,...
'   no=Mammoth,Pterodactyl
' Returns how many custom categories were loaded.
'--------------------------------------------------------------
Public Function LoadPacks(ByVal folder As String) As Long
    Dim fn As String, files() As String, n As Long, i As Long
    On Error GoTo Done
    fn = Dir$(folder & "*.txt")
    Do While Len(fn) > 0
        ReDim Preserve files(0 To n)
        files(n) = fn
        n = n + 1
        fn = Dir$
    Loop
    For i = 0 To n - 1
        LoadPackFile folder & files(i)
    Next
Done:
    LoadPacks = mNumCats - mBuiltin
End Function

Private Sub LoadPackFile(ByVal path As String)
    Dim f As Integer, ln As String, p As Long, k As String, v As String
    Dim t As String, isP As String, clue As String, yes As String, no As String
    On Error GoTo Fail
    f = FreeFile
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)
        If Len(ln) = 0 Then
            ' blank line
        ElseIf Left$(ln, 1) = ";" Or Left$(ln, 1) = "#" Then
            ' comment
        ElseIf Left$(ln, 1) = "[" Then
            CommitPack t, isP, clue, yes, no
            p = InStr(ln, "]")
            If p > 2 Then t = Trim$(Mid$(ln, 2, p - 2)) Else t = ""
            isP = "": clue = "": yes = "": no = ""
        Else
            p = InStr(ln, "=")
            If p > 1 Then
                k = LCase$(Trim$(Left$(ln, p - 1)))
                v = Trim$(Mid$(ln, p + 1))
                Select Case k
                Case "is"
                    isP = v
                Case "clue"
                    clue = v
                Case "yes"
                    If Len(yes) > 0 Then yes = yes & "," & v Else yes = v
                Case "no"
                    If Len(no) > 0 Then no = no & "," & v Else no = v
                End Select
            End If
        End If
    Loop
    CommitPack t, isP, clue, yes, no
    Close #f
    Exit Sub
Fail:
    On Error Resume Next
    If f <> 0 Then Close #f
End Sub

Private Sub CommitPack(ByVal t As String, ByVal isP As String, ByVal clue As String, ByVal yes As String, ByVal no As String)
    Dim parts() As String, i As Long, n As Long
    If Len(t) = 0 Or Len(yes) = 0 Then Exit Sub
    parts = Split(yes, ",")
    For i = 0 To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then n = n + 1
    Next
    If n < 4 Then Exit Sub      ' need at least 4 right answers
    If Len(isP) = 0 Then isP = "in the " & t & " group"
    AddCat t, isP, yes, no, clue
End Sub

Public Function CategoryCount() As Long
    CategoryCount = mNumCats
End Function

Public Sub ResetRuleHistory()
    Dim i As Long
    For i = 0 To MAX_CATS - 1
        mCatUsed(i) = False
    Next
    mLastKind = -1
End Sub

Public Function RandRange(ByVal lo As Long, ByVal hi As Long) As Long
    If hi <= lo Then
        RandRange = lo
    Else
        RandRange = lo + Int(Rnd * (hi - lo + 1))
    End If
End Function

Public Function CategorySize(ByVal cat As Long) As Long
    CategorySize = mCatLen(cat)
End Function

Public Function ModeName(ByVal gm As Long) As String
    Select Case gm
    Case MODE_MATH: ModeName = "Math"
    Case MODE_WORDS: ModeName = "Words"
    Case Else: ModeName = "Mixed"
    End Select
End Function

'--------------------------------------------------------------
' Build the rule for a level (lvl = rule difficulty, 1 and up)
'--------------------------------------------------------------
Public Function MakeRule(ByVal lvl As Long, ByVal gameMode As Long) As GameRule
    Dim r As GameRule
    Dim useWords As Boolean

    If lvl < 1 Then lvl = 1
    If gameMode = MODE_WORDS Then
        useWords = True
    ElseIf gameMode = MODE_MIXED Then
        useWords = ((lvl Mod 2) = 0)
    End If

    If useWords Then
        r.Kind = RK_CATEGORY
        r.Cat = PickCategory()
        r.Title = mCatTitle(r.Cat)
    Else
        r.Kind = PickMathKind(lvl)
        Select Case r.Kind
        Case RK_MULTIPLES
            If lvl <= 2 Then
                r.N = PickFrom("2,5,10")
            ElseIf lvl <= 5 Then
                r.N = PickFrom("3,4,5,6")
            ElseIf lvl <= 8 Then
                r.N = PickFrom("4,6,7,8,9")
            Else
                r.N = PickFrom("6,7,8,9,11,12")
            End If
            r.MaxV = r.N * 12
            If r.MaxV > 99 Then r.MaxV = 99
            If r.MaxV < 30 Then r.MaxV = 30
            r.Title = "Multiples of " & r.N

        Case RK_FACTORS
            If lvl <= 6 Then
                r.N = PickFrom("12,16,18,20,24,30")
            Else
                r.N = PickFrom("24,30,36,40,42,48,56,60,72")
            End If
            r.MaxV = r.N + 10
            r.Title = "Factors of " & r.N

        Case RK_PRIMES
            r.MaxV = 30 + (lvl - 5) * 10
            If r.MaxV < 30 Then r.MaxV = 30
            If r.MaxV > 99 Then r.MaxV = 99
            r.Title = "Prime Numbers"

        Case RK_EVENODD
            r.N = RandRange(0, 1)
            r.MaxV = 20 + lvl * 10
            If r.MaxV > 99 Then r.MaxV = 99
            If r.N = 0 Then r.Title = "Even Numbers" Else r.Title = "Odd Numbers"

        Case RK_COMPARE
            r.Ops = RandRange(0, 1)
            r.N = RandRange(3, 7) * 10
            If lvl >= 3 Then r.N = r.N + RandRange(0, 9)
            r.MaxV = 99
            If r.Ops = 0 Then r.Title = "Less than " & r.N Else r.Title = "Greater than " & r.N

        Case RK_SQUARES
            If lvl >= 9 Then r.MaxV = 144 Else r.MaxV = 100
            r.Title = "Square Numbers"

        Case Else
            If lvl <= 2 Then
                r.N = RandRange(5, 10): r.Ops = 1
            ElseIf lvl <= 4 Then
                r.N = RandRange(6, 15): r.Ops = 2
            ElseIf lvl <= 7 Then
                r.N = RandRange(8, 20): r.Ops = 3
            Else
                r.N = RandRange(10, 30): r.Ops = 4
            End If
            r.Title = "Equals " & r.N
        End Select
    End If

    mLastKind = r.Kind
    MakeRule = r
End Function

Private Function PickMathKind(ByVal lvl As Long) As Long
    Dim cand(0 To 7) As Long, n As Long, k As Long, tries As Long
    cand(0) = RK_MULTIPLES
    cand(1) = RK_EQUALS
    n = 2
    If lvl <= 4 Then
        cand(n) = RK_EVENODD: n = n + 1
        cand(n) = RK_COMPARE: n = n + 1
    End If
    If lvl >= 3 Then
        cand(n) = RK_FACTORS: n = n + 1
    End If
    If lvl >= 5 Then
        cand(n) = RK_PRIMES: n = n + 1
        cand(n) = RK_SQUARES: n = n + 1
    End If
    For tries = 1 To 8
        k = cand(RandRange(0, n - 1))
        If k <> mLastKind Then Exit For
    Next
    PickMathKind = k
End Function

Private Function PickCategory() As Long
    Dim i As Long, freeCount As Long, pick As Long
    For i = 0 To mNumCats - 1
        If Not mCatUsed(i) Then freeCount = freeCount + 1
    Next
    If freeCount = 0 Then
        For i = 0 To mNumCats - 1
            mCatUsed(i) = False
        Next
        freeCount = mNumCats
    End If
    pick = RandRange(1, freeCount)
    For i = 0 To mNumCats - 1
        If Not mCatUsed(i) Then
            pick = pick - 1
            If pick = 0 Then
                mCatUsed(i) = True
                PickCategory = i
                Exit Function
            End If
        End If
    Next
End Function

Private Function PickFrom(ByVal numList As String) As Long
    Dim parts() As String
    parts = Split(numList, ",")
    PickFrom = CLng(Val(parts(RandRange(0, UBound(parts)))))
End Function

'--------------------------------------------------------------
' Fill one square. good=True makes something that fits the rule.
'--------------------------------------------------------------
Public Sub MakeCell(r As GameRule, ByVal good As Boolean, c As GridCell)
    Dim v As Long, target As Long
    c.Filled = True
    c.Good = good

    Select Case r.Kind
    Case RK_MULTIPLES
        If good Then
            v = r.N * RandRange(1, r.MaxV \ r.N)
        Else
            Do
                v = RandRange(1, r.MaxV)
            Loop While (v Mod r.N) = 0
        End If

    Case RK_FACTORS
        If good Then
            Do
                v = RandRange(1, r.N)
            Loop While (r.N Mod v) <> 0
        Else
            Do
                v = RandRange(2, r.MaxV)
            Loop While (r.N Mod v) = 0
        End If

    Case RK_PRIMES
        If good Then
            Do
                v = RandRange(2, r.MaxV)
            Loop Until IsPrime(v)
        Else
            Do
                v = RandRange(1, r.MaxV)
            Loop While IsPrime(v)
        End If

    Case RK_EVENODD
        Do
            v = RandRange(1, r.MaxV)
        Loop While ((v Mod 2) = r.N) <> good

    Case RK_COMPARE
        If r.Ops = 0 Then
            If good Then v = RandRange(1, r.N - 1) Else v = RandRange(r.N, 99)
        Else
            If good Then v = RandRange(r.N + 1, 99) Else v = RandRange(1, r.N)
        End If

    Case RK_SQUARES
        If good Then
            v = RandRange(1, Int(Sqr(r.MaxV) + 0.5))
            v = v * v
        Else
            Do
                v = RandRange(2, r.MaxV)
            Loop While IsSquare(v)
        End If

    Case RK_EQUALS
        If good Then
            target = r.N
        Else
            Do
                target = r.N + RandRange(-3, 3)
            Loop While target = r.N Or target < 1
        End If
        c.Value = target
        c.Txt = MakeExpr(target, r.Ops)
        Exit Sub

    Case Else
        c.Value = 0
        If good Then
            c.Txt = mItems(mCatStart(r.Cat) + RandRange(0, mCatLen(r.Cat) - 1))
        Else
            c.Txt = RandomDistractor(r.Cat)
        End If
        Exit Sub
    End Select

    c.Value = v
    c.Txt = CStr(v)
End Sub

Private Function IsMember(ByVal cat As Long, ByVal w As String) As Boolean
    Dim i As Long
    For i = mCatStart(cat) To mCatStart(cat) + mCatLen(cat) - 1
        If StrComp(mItems(i), w, vbTextCompare) = 0 Then
            IsMember = True
            Exit Function
        End If
    Next
End Function

' A wrong answer for this category. Never returns a word that is
' also in the category (custom packs may overlap built-in lists).
Private Function RandomDistractor(ByVal cat As Long) As String
    Dim c2 As Long, w As String, tries As Long
    For tries = 1 To 40
        If mTrickLen(cat) > 0 And Rnd < 0.25 Then
            w = mItems(mTrickStart(cat) + RandRange(0, mTrickLen(cat) - 1))
        Else
            Do
                c2 = RandRange(0, mNumCats - 1)
            Loop While c2 = cat
            w = mItems(mCatStart(c2) + RandRange(0, mCatLen(c2) - 1))
        End If
        If Not IsMember(cat, w) Then Exit For
    Next
    RandomDistractor = w
End Function

' Builds a little sum whose answer is t, e.g. "4+5", "12-3", "3x3", "18/2"
Private Function MakeExpr(ByVal t As Long, ByVal ops As Long) As String
    Dim a As Long, b As Long, tries As Long
    For tries = 1 To 10
        Select Case RandRange(1, ops)
        Case 1
            If t >= 2 Then a = RandRange(1, t - 1) Else a = 0
            MakeExpr = a & "+" & (t - a)
            Exit Function
        Case 2
            b = RandRange(1, 9)
            MakeExpr = (t + b) & "-" & b
            Exit Function
        Case 3
            a = RandomDivisor(t)
            If a > 0 Then
                MakeExpr = a & Chr$(215) & (t \ a)
                Exit Function
            End If
        Case Else
            b = RandRange(2, 5)
            MakeExpr = (t * b) & Chr$(247) & b
            Exit Function
        End Select
    Next
    a = RandRange(0, t)
    MakeExpr = a & "+" & (t - a)
End Function

Private Function RandomDivisor(ByVal t As Long) As Long
    Dim d As Long, cnt As Long, pick As Long
    For d = 2 To t \ 2
        If (t Mod d) = 0 Then cnt = cnt + 1
    Next
    If cnt = 0 Then Exit Function
    pick = RandRange(1, cnt)
    cnt = 0
    For d = 2 To t \ 2
        If (t Mod d) = 0 Then
            cnt = cnt + 1
            If cnt = pick Then
                RandomDivisor = d
                Exit Function
            End If
        End If
    Next
End Function

Public Function IsPrime(ByVal v As Long) As Boolean
    Dim d As Long
    If v < 2 Then Exit Function
    For d = 2 To Int(Sqr(v))
        If (v Mod d) = 0 Then Exit Function
    Next
    IsPrime = True
End Function

Private Function IsSquare(ByVal v As Long) As Boolean
    Dim k As Long
    k = Int(Sqr(v) + 0.5)
    IsSquare = (k * k = v)
End Function

Private Function SmallestFactor(ByVal v As Long) As Long
    Dim d As Long
    For d = 2 To v
        If (v Mod d) = 0 Then
            SmallestFactor = d
            Exit Function
        End If
    Next
    SmallestFactor = v
End Function

'--------------------------------------------------------------
' Text helpers
'--------------------------------------------------------------
Public Function WrongMessage(r As GameRule, c As GridCell) As String
    Dim d As Long
    Select Case r.Kind
    Case RK_MULTIPLES
        WrongMessage = c.Txt & " is not a multiple of " & r.N & "."
    Case RK_FACTORS
        WrongMessage = c.Txt & " is not a factor of " & r.N & "."
    Case RK_PRIMES
        If c.Value < 2 Then
            WrongMessage = "1 is not a prime number."
        Else
            d = SmallestFactor(c.Value)
            WrongMessage = c.Txt & " is not prime:  " & d & " " & Chr$(215) & " " & (c.Value \ d) & " = " & c.Value & "."
        End If
    Case RK_EVENODD
        If r.N = 0 Then
            WrongMessage = c.Txt & " is odd, not even."
        Else
            WrongMessage = c.Txt & " is even, not odd."
        End If
    Case RK_COMPARE
        If r.Ops = 0 Then
            WrongMessage = c.Txt & " is not less than " & r.N & "."
        Else
            WrongMessage = c.Txt & " is not greater than " & r.N & "."
        End If
    Case RK_SQUARES
        WrongMessage = c.Txt & " is not a square number."
    Case RK_EQUALS
        WrongMessage = c.Txt & " = " & c.Value & ",  not " & r.N & "."
    Case Else
        WrongMessage = c.Txt & " is not " & mCatOne(r.Cat) & "."
    End Select
End Function

Public Function RuleHint(r As GameRule) As String
    Select Case r.Kind
    Case RK_MULTIPLES
        RuleHint = "Count by " & r.N & "s:  " & r.N & ", " & (r.N * 2) & ", " & (r.N * 3) & ", " & (r.N * 4) & " ..."
    Case RK_FACTORS
        RuleHint = "Numbers that divide into " & r.N & " with nothing left over."
    Case RK_PRIMES
        RuleHint = "A prime can only be divided evenly by 1 and itself."
    Case RK_EVENODD
        If r.N = 0 Then
            RuleHint = "Even numbers end in 0, 2, 4, 6 or 8."
        Else
            RuleHint = "Odd numbers end in 1, 3, 5, 7 or 9."
        End If
    Case RK_COMPARE
        If r.Ops = 0 Then
            RuleHint = "Gobble every number smaller than " & r.N & "."
        Else
            RuleHint = "Gobble every number bigger than " & r.N & "."
        End If
    Case RK_SQUARES
        RuleHint = "A number times itself:  1, 4, 9, 16, 25 ..."
    Case RK_EQUALS
        RuleHint = "Gobble every problem whose answer is " & r.N & "."
    Case Else
        If Len(mCatClue(r.Cat)) > 0 Then
            RuleHint = mCatClue(r.Cat)
        Else
            RuleHint = "Gobble every word that is " & mCatOne(r.Cat) & "."
        End If
    End Select
End Function
