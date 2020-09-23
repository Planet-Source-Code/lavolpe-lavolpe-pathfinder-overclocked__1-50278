VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPathFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' 60% rewrite from previous posting -- new ideas!.
' The logic/rationale I'm using can be read in the HowItWorks.RTF file provided
' I have included many remarks in the code to help follow along

' Major improvement: Now 100% capability of finding the shortest path
' Performance hit: increase time required; but still kept to an average of < 1 second

' If you should find, while playing with this, a scenario where the shortest
' path was not found. Please take a screen capture of the A/Z placement and
' please send it to me. But first, try upping the aLimit & zLimit variables
' in the routine JoinPaths. I currently have these set to abort any path legs
' that are longer than the completed segment (a/z) + 20%. Up it to 30/50% and
' try again. The limiting factors reduce the number of new legs that are spawned

' I've ran through 100's of scenarios, including ones that
' I knew failed in the previous version, and could not make this version
' return anything less than the shortest path with vAccuracy set @ 100%

' Next tasks include....
' 1. Allow cross-over paths (like a circuit board) where a horizontal leg can
'   cross over a vertical leg & vice versa. Needed for my project @ work
'   Hmmm? This is the next logical step for a 3D mapping routine.
' 2. Allow multiple path limiters (gaming purposes). Example. faster to
'    move downhill than uphill.
' 3. Include "smart" diagonal paths. This is a lower priority of mine
' 4. Maybe include capability to dictate some paths as one-way only.
'    Currently all paths are multidirectional.

' used for debugging
Private Declare Function RedrawWindow Lib "user32.dll" (ByVal hwnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
' routines only used for testing/debugging & can be removed
'DisplayPartialPath
'DisplayRectangle
'AnimatePts


' Primary APIs used to calculate paths
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long
Private Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Private Const RGN_OR As Long = 2
Private Const RGN_XOR As Long = 3

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type CompletePathStructure
    Length As Long
    aIndex As Long
    zIndex As Long
    aRestrictor As Long
    zRestrictor As Long
    Accuracy As Byte
End Type
Private Type SegmentStructure
    Parent As Long
    Length As Long
    X As Integer
    Y As Integer
    RectID As Integer
    Status As Byte
End Type
Private Type NodeStructure
    LinkIndex As Integer    ' Index into the Links() array, if any
    CollideIndex As Integer ' Index into the CollisionPts() array
    Heuristics(0 To 1) As Long ' relative distance from A/Z node
    Owner As Byte              ' Visited by A/Z node (0=A, 1=Z)
End Type
Private Type LINKSTRUCTURE
    ' array of adjacent rectangles by NodeRects() index
    Adjacents() As Integer
End Type

Private NodeRects() As RECT        'Array of rectangles in the passed region
Private Legs() As NodeStructure    'Index is reference to NodeRect array (1:1 ratio)
Private Links() As LINKSTRUCTURE   'List of adjacent rectangles
Private CollisionPts() As Long     'Created on-demand & used during path finding only
Private vPaths() As SegmentStructure 'Primary array used during pathfinding only
Private vTempA() As Long           ' used for stacking A/Z path legs to explore
Private vTempZ() As Long
Private FullPath As CompletePathStructure ' data pertaining to shortest path

Private Sub Class_Terminate()
' remove all arrays
' Note: vTempA & vTempZ are removed in CreatePaths routine
Erase vPaths        ' may be removed in CreatePaths, may not be
Erase CollisionPts
Erase Links
Erase Legs
Erase NodeRects
End Sub

Public Sub SetMapRegion(ByVal hRgn As Long, Optional ByVal sizeRestriction As Byte = 0)
' primary entry point into the class
' The user passes the windows region containing their map & an optional restrictor value
' to restrict the size of the paths. This restrictor is sizeRestriction x sizeRestriction
If sizeRestriction Then
    If sizeRestriction = 255 Then sizeRestriction = 254 ' prevent error
    ' ensure an even number is passed
    If RestrictPaths(hRgn, sizeRestriction + (sizeRestriction Mod 2)) = False Then Exit Sub
Else
    If ExtractRectangles(hRgn) = False Then Exit Sub
End If
' call function to build the link table
BuildLinkTable
End Sub


Private Function RestrictPaths(ByVal hRgn As Long, ByVal rSize As Byte) As Long

On Error GoTo EH    ' troubleshooting

' I can think of a few ways to restrict paths to a specific minimum size
' This is one way that is fast

' Note: Function is called only when the map is assigned/reassigned to this class

If rSize = 0 Then
    ' no need to rebuild the map if rSize = 0
    RestrictPaths = ExtractRectangles(hRgn)
    Exit Function
End If

Dim invRgn As Long, tRgn As Long
Dim iRect As RECT, tRect As RECT
Dim X As Long, Y As Long, Z As Long

' create the inverse of the map region so all we have are the obstacles
GetRgnBox hRgn, iRect
invRgn = CreateRectRgnIndirect(iRect)
CombineRgn invRgn, hRgn, invRgn, RGN_XOR

' get the rectangle structure of the inverse region
If ExtractRectangles(invRgn) = False Then
    DeleteObject invRgn
    Exit Function
End If

' loop thru twice
' 1st loop to increase obstacles 1/2 of the Bot size in all directions
For X = 0 To UBound(NodeRects)
    ' add the inflated rectangle back to the obstacle region
    InflateRect NodeRects(X), rSize \ 2, rSize \ 2
    tRgn = CreateRectRgnIndirect(NodeRects(X))
    CombineRgn invRgn, invRgn, tRgn, RGN_OR
    DeleteObject tRgn
Next
' now we add the top & bottom border to the obstacle region
SetRect tRect, iRect.Left, iRect.Top, iRect.Right, iRect.Top + rSize \ 2
tRgn = CreateRectRgnIndirect(tRect)
CombineRgn invRgn, invRgn, tRgn, RGN_OR
DeleteObject tRgn

SetRect tRect, iRect.Left, iRect.Bottom - rSize \ 2, iRect.Right, iRect.Bottom
tRgn = CreateRectRgnIndirect(tRect)
CombineRgn invRgn, invRgn, tRgn, RGN_OR
DeleteObject tRgn

' now add the left & right borders to the obstacle region
SetRect tRect, iRect.Right - rSize \ 2, iRect.Top, iRect.Right, iRect.Bottom
tRgn = CreateRectRgnIndirect(tRect)
CombineRgn invRgn, invRgn, tRgn, RGN_OR
DeleteObject tRgn

SetRect tRect, iRect.Left, iRect.Top, iRect.Left + rSize \ 2, iRect.Bottom
tRgn = CreateRectRgnIndirect(tRect)
CombineRgn invRgn, invRgn, tRgn, RGN_OR
DeleteObject tRgn
   
' 2nd Loop
' loop thru to find any intersections between nodes vertically
' that result in a path too small for the bot to traverse
' Intersections between nodes horizontally that are too small
' are automatically removed simply by inflating the rectangles
' as we did in the above step
ExtractRectangles invRgn

For X = 0 To UBound(NodeRects) - 1
    
    Z = NodeRects(X).Bottom
    NodeRects(X).Bottom = NodeRects(X).Bottom + rSize
    For Y = X + 1 To UBound(NodeRects)
    
        If IntersectRect(iRect, NodeRects(X), NodeRects(Y)) Then
        
            ' add the intersection rectangle to the obstacle region
            tRgn = CreateRectRgnIndirect(iRect)
            CombineRgn invRgn, invRgn, tRgn, RGN_OR
            DeleteObject tRgn
            
        End If
        If NodeRects(Y).Top > Z Then Exit For
    Next
    
Next

' get the overall size of the original region, once again
GetRgnBox hRgn, iRect
' create a blank region the size of the original region
tRgn = CreateRectRgnIndirect(iRect)
' now we get the inverse of the obstacle region, which will
' now be the source region with 'expanded' obstacles to restrict paths
CombineRgn tRgn, invRgn, tRgn, RGN_XOR

' debugging purposes
'X = CreateSolidBrush(vbBlue)
'FillRgn frmMain.picWork.hdc, invRgn, X
'DeleteObject X

DeleteObject invRgn

' debugging purposes
'X = CreateSolidBrush(vbWhite)
'FillRgn frmMain.picWork.hdc, tRgn, X
'DeleteObject X

' last time extracting the rectangles
RestrictPaths = ExtractRectangles(tRgn)
DeleteObject tRgn

Exit Function
EH:
Stop
Resume
End Function


Private Sub BuildLinkTable()

' Function will build a link table identifying adjacent nodes to each node in the map
' Note: Function is called only when the map is assigned/reassigned to this class

On Error GoTo EH        ' debugging purposes

Dim X As Long, Y As Long
Dim iRect As RECT, aRect As RECT
Dim I As Long

' initialize the Links array.
ReDim Links(0)
ReDim Links(0).Adjacents(0 To 0)
ReDim Legs(0 To UBound(NodeRects))

' begin determining adjacent nodes
For X = 0 To UBound(NodeRects) - 1
    ' need to extend the bottom edge 1 pixel to force an intersection between
    ' this node's bottom edge & those rectangles sharing the edge as their top value
    SetRect aRect, NodeRects(X).Left, NodeRects(X).Top, NodeRects(X).Right, NodeRects(X).Bottom + 1
    
    For Y = X + 1 To UBound(NodeRects)
    
        If IntersectRect(iRect, aRect, NodeRects(Y)) Then   ' we have a link
                
                ' this routine will merge rectangles that have the same
                ' left/right values & are vertically adjacent.
                ' By merging them, we can drastically reduce the number of
                ' rectangles that can be considered in the path routines
                ' The bottom line, is that merging them reduces path calculation
                ' time significantly
                
                If aRect.Right = NodeRects(Y).Right And aRect.Left = NodeRects(Y).Left Then
                    ' extend the test rectangle's bottom to the
                    ' adjacent rectangle's bottom value + 1
                    aRect.Bottom = NodeRects(Y).Bottom + 1
                    ' extend the current rectangle's bottom to the
                    ' adjacent rectangle's bottom value
                    NodeRects(X).Bottom = NodeRects(Y).Bottom
                    ' by making the bottom value less than the top value,
                    ' we make the rectangle APIs refuse this rectangle
                    ' in any of their calculations. Ideally, we would
                    ' simply remove the rectangle from the array, but the
                    ' tNodes() array is a one-to-one reference to this array &
                    ' removing the rectangles will corrupt the link table
                    NodeRects(Y).Bottom = NodeRects(Y).Top - 1
                    I = I + 1   ' debugging only
                Else
                    CreateNewLink X, Y
                End If
        Else
            ' since window regions are designed as stacked rectangles that do not overlap
            ' determining when to stop checking for adjacent nodes is simple...
            If NodeRects(Y).Top > NodeRects(X).Bottom Then Exit For
        End If
    Next
Next

'Debug.Print UBound(NodeRects); " rectangles were reduced to "; UBound(NodeRects) - I, I
Exit Sub

EH:
Stop
Resume
End Sub

Private Sub CreateNewLink(ByVal rIndex As Integer, ByVal aIndex As Integer)
' extension of the BuildLink routine
' This will fill the Links() array

Dim nrLinks As Integer, myIndex As Integer
On Error Resume Next   ' handle uninitialized arrays
' Build the Aside Link first & can trigger uninitialized array errors
If Legs(rIndex).LinkIndex = 0 Then
    myIndex = UBound(Links)
    myIndex = myIndex + 1
    ReDim Preserve Links(0 To myIndex)
    Legs(rIndex).LinkIndex = myIndex
Else
    myIndex = Legs(rIndex).LinkIndex
End If
nrLinks = UBound(Links(myIndex).Adjacents)
nrLinks = nrLinks + 1
ReDim Preserve Links(myIndex).Adjacents(1 To nrLinks)
Links(myIndex).Adjacents(nrLinks) = aIndex
If Err Then Err.Clear

' Now build the ZSide Link -- won't trigger any errors
' Since any uninitialized arrays were initialized above
If Legs(aIndex).LinkIndex = 0 Then
    myIndex = UBound(Links) + 1
    ReDim Preserve Links(0 To myIndex)
    ReDim Links(myIndex).Adjacents(1 To 1)
    Links(myIndex).Adjacents(1) = rIndex
    Legs(aIndex).LinkIndex = myIndex
Else
    myIndex = Legs(aIndex).LinkIndex
    For nrLinks = 1 To UBound(Links(myIndex).Adjacents)
        If Links(myIndex).Adjacents(nrLinks) = rIndex Then Exit For
    Next
    If nrLinks > UBound(Links(myIndex).Adjacents) Then
        ReDim Preserve Links(myIndex).Adjacents(1 To nrLinks)
        Links(myIndex).Adjacents(nrLinks) = rIndex
    End If
End If
End Sub

Private Sub CreateCollisionPts(ByVal rIndex As Integer)
' On demand array. This little baby does a lot of good....
' Since the map is basically a huge array of 1 pixel cells, we
' have no real way of determining when a path leg is traveling the
' same path as a previous leg. However, since obstacles force the
' paths to change directions, we can at least check for duplicate
' paths whenever a path collides with an obstacle. This way, we
' significantly reduce the number of path legs that are generated.

If Legs(rIndex).CollideIndex Then Exit Sub
Dim myIndex As Integer
On Error Resume Next   ' handle uninitialized arrays
myIndex = UBound(CollisionPts, 2)
myIndex = myIndex + 1
ReDim Preserve CollisionPts(0 To 3, 1 To myIndex)
'-- 0=top left, 1=bot left, 2=top right, 3=bot right
Legs(rIndex).CollideIndex = myIndex
If Err Then Err.Clear
End Sub

Public Function CreatePaths(StartPtX As Long, StartPtY As Long, _
                EndPtX As Long, EndPtY As Long, vRtnArray() As Long, _
                Optional vAccuracyPct As Byte = 100) As Long

Dim myCounter As Long   ' for debugging & testing
Dim bShowWork As Boolean ' for debugging & testing
bShowWork = frmMain.chkShowWork.Enabled And frmMain.chkShowWork = 1
' remove all refs to bShowWork & myCounter in this routine if you use this
' class in your own project. Also rem out/remove the routine AnimatePts.

'1. Get A & Z starting Rectangles & A/Z XY coords &
'2. Test to see if A & Z are in the same rectangle?
Dim StartRect As Integer, EndRect As Integer
Dim Looper As Integer

StartRect = -1
EndRect = -1
' locate the rectangles containing the start & stop nodes
Do
    If StartRect < 0 Then
        If PtInRect(NodeRects(Looper), StartPtX, StartPtY) Then StartRect = Looper
    End If
    If EndRect < 0 Then
        If PtInRect(NodeRects(Looper), EndPtX, EndPtY) Then EndRect = Looper
    End If
    Looper = Looper + 1
Loop Until (StartRect > 0 And EndRect > 0) Or Looper > UBound(Legs)
If StartRect < 0 Or EndRect < 0 Then
    ' error nodea or nodez is not in the map or is in an enclosed region
    ' If a restrictor on path size was used, then the node may overlap an area
    ' which prevents travel from NodeA to NodeZ
    Exit Function
End If
If StartRect = EndRect Then
    ' direct route between nodes; draw straight edge
    MsgBox "Straight path -- no fun :("
    ReDim vRtnArray(0 To 1)
    vRtnArray(0) = MakeLong(StartPtX + 0, StartPtY + 0)
    vRtnArray(1) = MakeLong(EndPtX + 0, EndPtY + 0)
    CreatePaths = Abs(StartPtX - EndPtX) + Abs(StartPtY - EndPtY)
    Exit Function
End If

'3. Prime A & Z Paths
Erase CollisionPts
CalculateHeuristics StartPtX, StartPtY, EndPtX, EndPtY

' dimension array for initial A/Z path legs & the zero-member
ReDim vPaths(0 To 1)
FullPath.Accuracy = vAccuracyPct
If FullPath.Accuracy > 100 Then FullPath.Accuracy = 100

' used simply as a list of active legs
ReDim vTempA(0 To 1, 0 To 0)
ReDim vTempZ(0 To 1, 0 To 1)
' element 0 = index into the vPaths() array
' element 1 = sort value

' initialize the A & Z paths (1=A, 2=Z)
vPaths(0).Parent = -1
vPaths(0).X = StartPtX
vPaths(0).Y = StartPtY
vPaths(0).RectID = StartRect
vPaths(1).Parent = -1
vPaths(1).X = EndPtX
vPaths(1).Y = EndPtY
vPaths(1).RectID = EndRect
' prime the temp arrays
vTempZ(0, 1) = 1 'cause of the way following DO loop runs, the
vTempZ(1, 1) = 1 '1st Z array item will be stripped. So we do it twice here

Dim bCanContinue(0 To 1) As Boolean
Dim linkIdx As Integer
Dim pIndex As Long

'4. Generate Paths
bCanContinue(0) = True
Do
    For Looper = 0 To 1
        If bCanContinue(Looper) Then
            ' get the vPath() index to expand & reset boolean value
            pIndex = Choose(Looper + 1, vTempA(0, 0), vTempZ(0, 0))
            bCanContinue(Looper) = False    ' reset for now
            
            ' terminate long, unfinished paths early if possible
            If Looper Then
                If vPaths(pIndex).Length < FullPath.zRestrictor Or FullPath.Length = 0 Then bCanContinue(Looper) = True
            Else
                If vPaths(pIndex).Length < FullPath.aRestrictor Or FullPath.Length = 0 Then bCanContinue(Looper) = True
            End If
            If bCanContinue(Looper) Then
            
                ' reset. tells routine if more temp paths are ready for exploring
                bCanContinue(Looper) = False
                
                ' for each adjacent rectangle, process it
                For linkIdx = 1 To UBound(Links(Legs(vPaths(pIndex).RectID).LinkIndex).Adjacents)
                    ' call function to do all the work & return a result
                    Select Case MoveToRect(pIndex, Links(Legs(vPaths(pIndex).RectID).LinkIndex).Adjacents(linkIdx), Looper + 1)
                    Case 2  ' found a path
                        If bShowWork Then AnimatePts    ' remove when final
                        JoinPaths UBound(vPaths), Looper
                        ' if 1st path found requested then exit now
                        If FullPath.Accuracy = 0 Then Exit Do
                        bCanContinue(Looper) = True
                    Case 1  ' valid next leg
                        If bShowWork Then AnimatePts    ' remove when final
                        AddToTempArray UBound(vPaths), Looper
                        bCanContinue(Looper) = True
                    Case 0         ' backtracking or can't move
                    Case Else
                        bCanContinue(Looper) = True
                    End Select
                Next linkIdx
            Else    ' terminating 'cause partial path is too long & a shorter complete path found
                myCounter = myCounter + 1           ' debugging/testing
                vPaths(pIndex).Status = 64 Or Looper   ' mark as too long
            End If
            ' if no legs could be spawned, then this was a dead end -- mark it as such
            If bCanContinue(Looper) = False And vPaths(pIndex).Status < 64 Then vPaths(pIndex).Status = 128 Or Looper
        End If
        ' sort vFinder on either the A or Z side to get next path leg to try
        If RebuildTempArray(Abs(Looper - 1)) Then
            ' set flag to continue processing path legs
            bCanContinue(Abs(Looper - 1)) = True
        Else
            ' since we are doing bidirectional searches, if we ran out of
            ' Z-side nodes to expand, didn't find a path, and just finished
            ' the A-Side's turn, then there is no possible path or vice versa --
            ' no need to explore every available node
            If (UBound(vTempZ, 2) = 0 And FullPath.Length = 0 And Looper = 0) Then
                Exit Do
            Else
                If (UBound(vTempA, 2) = 0 And FullPath.Length = 0 And Looper = 1) Then
                    Exit Do
                Else
                    ' When totally out of nodes, we abort the Do Loop
                    If (UBound(vTempA, 2) = 0 And UBound(vTempZ, 2) = 0) Then Exit Do
                End If
            End If
            bCanContinue(Abs(Looper - 1)) = False
        End If
    Next Looper
Loop
' all done... return a completed path
If FullPath.Length Then
'    DisplayPartialPath FullPath.aIndex   ' debugging
'    DisplayPartialPath fullpath.zindex    ' debugging
    Debug.Print "path length chosen is "; FullPath.Length, UBound(vPaths) & " spawns", myCounter
    BuildReturnArray vRtnArray()
    ' clear the FullPath UDT & return the length
    With FullPath
        CreatePaths = .Length
        .Length = 0
        .aRestrictor = 0
        .zRestrictor = 0
        .aIndex = 0
        .zIndex = 0
        .Accuracy = 0
    End With
End If
Erase vTempA
Erase vTempZ
Erase bCanContinue
Erase vPaths
End Function

Private Function MoveToRect(ByVal pIndex As Long, ByVal nextNode As Long, ByVal Owner As Integer) As Long
'VPATH ARRAY STRUCTURE
'element 0 = ArrayIndex that spawned this leg (Parent)
'element 1 = total length up to & including this leg
'element 2 = current XY position (Hi/Lo words)
'element 3 = ref to Rectangle at XY   (-1 indicates deadend)
'element 4 = status

' don't allow backtracking over the same path legs
If IsBacktracking(pIndex, nextNode) Then Exit Function

Dim X As Long, Y As Long, newX As Long, newY As Long
Dim adjLength As Long, iCorner As Integer, bPathFound As Boolean
Dim tRect As RECT

' get the current legs X,Y position & path length up to this point
X = vPaths(pIndex).X
Y = vPaths(pIndex).Y
adjLength = vPaths(pIndex).Length

' see if the path is now complete & set variable if so
If Legs(nextNode).Owner > 0 And Legs(nextNode).Owner <> Owner Then bPathFound = True
    
' determine the next X,Y segment
' get the new Y coordinate
If NodeRects(nextNode).Top < NodeRects(vPaths(pIndex).RectID).Top Then ' moving north
    newY = NodeRects(nextNode).Bottom
Else
    newY = NodeRects(nextNode).Top
End If

' get the new X coordinate, if needed
With NodeRects(nextNode)
    SetRect tRect, .Left, .Top, .Right + 1, .Bottom
End With
If PtInRect(tRect, X, tRect.Top) Then
    newX = X
    adjLength = adjLength + Abs(newY - Y)
Else
    ' when forced to move left/right due to an obstacle, need to check a few more things
    ' determine which corner of the target node we are moving to
    ' & set the newX coord & also update the new Length
    If NodeRects(nextNode).Right < X Then ' moving east
        newX = NodeRects(nextNode).Right
        iCorner = 2 ' right corners
    Else
        newX = NodeRects(nextNode).Left   ' left corners (iCorner=0)
    End If
    ' when moving south, we add 1 to corner value: 0=top left, 1=bot left, 2=top right, 3=bot right
    If NodeRects(nextNode).Top < NodeRects(vPaths(pIndex).RectID).Top Then iCorner = iCorner + 1  ' bottom corners
    adjLength = adjLength + (Abs(newX - X) + Abs(newY - Y))
    
    ' now we see if we've already been here, create a collision pts array if needed
    CreateCollisionPts nextNode
    If CollisionPts(iCorner, Legs(nextNode).CollideIndex) Then ' been to this exact point already
        ' if we just got here and am a longer path, then we'll abort now; otherwise we'll take over the longer path
        If vPaths(CollisionPts(iCorner, Legs(nextNode).CollideIndex)).Length > adjLength Then
        ' We can do 1 of 2 things:
        '   1-terminate the longer path and all of its spawned paths
        '   2-adopt & take over the longer path (and all of its spawns) & also adjust leg lengths as we adopt
        ' since the other path already spawned & this path would spawn same legs, we'll adopt the longer path
        ' either way, we are forced to find all the spawns to either adjust lengths or terminate
            AssimilatePath CollisionPts(iCorner, Legs(nextNode).CollideIndex), adjLength - vPaths(CollisionPts(iCorner, Legs(nextNode).CollideIndex)).Length, Owner
            ' now we take over the previous paths by rerouting the spawning parent to our current leg
            vPaths(CollisionPts(iCorner, Legs(nextNode).CollideIndex)).Parent = pIndex
        End If
        ' We'll abort 'cause we are a longer path or we took over a path.
        ' We won't spawn a new path from here regardless
        Exit Function
    End If
    ' update the collision corner with a ref back to our new path leg
    CollisionPts(iCorner, Legs(nextNode).CollideIndex) = UBound(vPaths) + 1
End If

' add the new path leg to the vPaths array
MoveToRect = 1 + Abs(bPathFound)
ReDim Preserve vPaths(0 To UBound(vPaths) + 1)
With vPaths(UBound(vPaths))
    .Parent = pIndex
    .Length = adjLength
    .X = newX
    .Y = newY
    .RectID = nextNode
    .Status = Owner - 1
    If Not bPathFound Then Legs(.RectID).Owner = Owner
End With
End Function

Private Function IsBacktracking(ByVal pIndex As Long, ByVal nextNode As Long) As Boolean
' simple routine & a bit timeconsuming. But we need to know if we are backtracking
' otherwise a path can go on forever

' Loop until the parent = the original parent (-1)
Do
    If vPaths(pIndex).RectID = nextNode Then
         IsBacktracking = True
         Exit Do
    End If
    pIndex = vPaths(pIndex).Parent
Loop Until pIndex < 0
End Function

Private Sub AssimilatePath(ByVal fromIndex As Long, ByVal lenDiff As Long, ByVal Owner As Integer)
' recursive routine to find all spawns from a specified parent: fromIndex
' This routine is only called when a path is replicating another path exactly
' and the new path is shorter. Instead of allowing the class to simply
' create new path segments from the shorter path, we will find all the
' spawns from the longer path, and adjust the overall lengths. When this
' routine finishes, the calling routine will then divert the source leg
' to its own index as the new parent.

' dead-end, backtracking path
If ((vPaths(fromIndex).Status And 128) = 128) Then Exit Sub

' change the length
vPaths(fromIndex).Length = vPaths(fromIndex).Length + lenDiff

If Legs(vPaths(fromIndex).RectID).Owner <> Owner Then
    ' if this leg was a path terminator, then we need to
    ' recalculate the complete path to see if it is shorter
    ' than the shortest path at the moment
    JoinPaths fromIndex, Owner - 1
    Exit Sub
End If

' recursive... keep variables at a minimum
Dim nextIdx As Long, bCanExit As Boolean
' Note about bCanExit
' Whenever a parent spawns new legs, those legs are always sequentially
' indexed. i.e., parent 50 can spawn 3 new paths: 77, 78, 79. Therefore
' if we found a child, then if the next leg in vPaths is not a child,
' we can abort the loop vs going to the very end of the vPaths array

For nextIdx = fromIndex + 1 To UBound(vPaths)
    If vPaths(nextIdx).Parent = fromIndex Then
        bCanExit = True
        If vPaths(nextIdx).Status < 128 Then
            If ((vPaths(nextIdx).Status And 64) = 64) Then
                ' path was aborted because it was too long
                ' Now it will be shorter, so we add it back into
                ' the active legs array for spawning
                vPaths(nextIdx).Status = Owner - 1
                AddToTempArray nextIdx, Owner - 1
            End If
            ' now we find its children, if any
            AssimilatePath nextIdx, lenDiff, Owner
        End If
    Else
        If bCanExit Then Exit For
    End If
Next
End Sub

Private Sub JoinPaths(ByVal pIndex As Long, ByVal lSide As Integer)
' Function joins the paths explored by the A & Z sides.
' If the completed path is shorter than an existing path,
' the function updates the vPaths() array appropriately

Dim JoinCost As Long
Dim X As Long, bReplace As Boolean
' The hard part is determining which A or Z side path segment to join to...
' We know we collided on a specific node, so by searching the
' the vpaths() array to find that node, we then find the correct
' paths to join to. There may be more than 1 possible path leg, so we
' need to compare all the possibilities & pick the combination with
' the shortest overall length

For X = pIndex - 1 To 1 Step -1
    ' find the colliding node in the current vpaths() array, ignore deadends
    If vPaths(X).RectID = vPaths(pIndex).RectID And vPaths(X).Status < 128 Then
        
        If lSide = 1 Then   ' collided on an A-Side node from the Z-side path
            ' make sure the owner of the node is the A-Side
            If ((vPaths(X).Status And 1) <> 1) Then
                ' calculate the length of joining the 2 paths
                JoinCost = Abs(vPaths(pIndex).X - vPaths(X).X)
                JoinCost = JoinCost + Abs(vPaths(pIndex).Y - vPaths(X).Y)
                JoinCost = JoinCost + vPaths(pIndex).Length + vPaths(X).Length
                ' compare to see if shorter than existing shortest path
                If JoinCost < FullPath.Length Or FullPath.Length = 0 Then
                    FullPath.aIndex = X
                    FullPath.zIndex = pIndex
                    FullPath.Length = JoinCost
                    bReplace = True
                End If
            End If
        Else    ' collided on a Z-Side node from the A-side path
            ' make sure the owner of the node is the A-Side
            If ((vPaths(X).Status And 1) = 1) Then
                ' calculate the length of joining the 2 paths
                JoinCost = Abs(vPaths(pIndex).X - vPaths(X).X)
                JoinCost = JoinCost + Abs(vPaths(pIndex).Y - vPaths(X).Y)
                JoinCost = JoinCost + vPaths(pIndex).Length + vPaths(X).Length
                ' compare to see if shorter than existing shortest path
                If JoinCost < FullPath.Length Or FullPath.Length = 0 Then
                    FullPath.aIndex = pIndex
                    FullPath.zIndex = X
                    FullPath.Length = JoinCost
                    bReplace = True
                End If
            End If
        End If
    End If
Next
If bReplace Then
'    Debug.Print "Found shorter path "; fullpath.length

    ' calculate a length limiting value for both the A/Z sides.
    ' Every path we can terminate early reduces the total number of
    ' spawns exponentially
    Dim aLimit As Integer, zLimit As Integer
    aLimit = vPaths(FullPath.aIndex).Length
    zLimit = vPaths(FullPath.zIndex).Length
    FullPath.aRestrictor = aLimit * (FullPath.Accuracy / 100) + (aLimit * 0.2)
    FullPath.zRestrictor = zLimit * (FullPath.Accuracy / 100) + (zLimit * 0.2)

    ' because of the way the XY coords are set, the colliding
    ' paths will never collide on the same point. Therefore, we
    ' need to add an extra point in the vPaths() array so
    ' they join properly
    ReDim Preserve vPaths(0 To UBound(vPaths) + 1)
    With vPaths(UBound(vPaths))
        .Parent = pIndex  ' set the parent
        .Status = 128   ' set as deadend
    End With
    ' depending on which side of the path collided
    ' set the new XY coords for the dummy array item
    ' and also update the 0 member of the vPaths() array
    With FullPath
        If .aIndex = pIndex Then
            vPaths(UBound(vPaths)).X = vPaths(.zIndex).X
            vPaths(UBound(vPaths)).Y = vPaths(.zIndex).Y
            .aIndex = UBound(vPaths)
        Else
            vPaths(UBound(vPaths)).X = vPaths(.aIndex).X
            vPaths(UBound(vPaths)).Y = vPaths(.aIndex).Y
            .zIndex = UBound(vPaths)
        End If
    End With
    vPaths(pIndex).Status = 32 Or lSide     ' mark as complete
Else    ' couldn't find join the path or the path was
        ' longer than the shortest found so far
    vPaths(pIndex).Status = 64 Or lSide ' too long
End If
End Sub

Private Sub AddToTempArray(ByVal pIndex As Long, ByVal Owner As Integer)
' function adds new spawns to the temp appropriate temp arrays
' By using the heuristics calculated when the region was processed and then
' adding the relative distance from the current path (pIndex) to the target
' rectangle, we can prioritize which new path leg gets expanded first.
' Again, the idea is to find paths the fastest possible to help weed out
' potentially longer paths & kill them before they can spawn

If Owner Then   ' Z-side spawns
    ReDim Preserve vTempZ(0 To 1, 0 To UBound(vTempZ, 2) + 1)
    vTempZ(0, UBound(vTempZ, 2)) = pIndex
    vTempZ(1, UBound(vTempZ, 2)) = Legs(vPaths(pIndex).RectID).Heuristics(0) + Abs(vPaths(pIndex).X - vPaths(1).X) + Abs(vPaths(pIndex).Y - vPaths(1).Y)
Else            ' A-side spawns
    ReDim Preserve vTempA(0 To 1, 0 To UBound(vTempA, 2) + 1)
    vTempA(0, UBound(vTempA, 2)) = pIndex
    vTempA(1, UBound(vTempA, 2)) = Legs(vPaths(pIndex).RectID).Heuristics(1) + Abs(vPaths(pIndex).X - vPaths(0).X) + Abs(vPaths(pIndex).Y - vPaths(0).Y)
End If
End Sub

Private Function RebuildTempArray(ByVal Owner As Integer) As Boolean
' function will shift the temp arrays to the left by one & then
' re-sort the array. This is called every time we finish expanding
' a new path leg
If Owner Then
    If UBound(vTempZ, 2) Then
        vTempZ(0, 0) = vTempZ(0, UBound(vTempZ, 2))
        vTempZ(1, 0) = vTempZ(1, UBound(vTempZ, 2))
        ReDim Preserve vTempZ(0 To 1, 0 To UBound(vTempZ, 2) - 1)
        RebuildTempArray = True
        TriQuickSortLong vTempZ, 1
    End If
Else
    If UBound(vTempA, 2) Then
        vTempA(0, 0) = vTempA(0, UBound(vTempA, 2))
        vTempA(1, 0) = vTempA(1, UBound(vTempA, 2))
        ReDim Preserve vTempA(0 To 1, 0 To UBound(vTempA, 2) - 1)
        RebuildTempArray = True
        TriQuickSortLong vTempA, 1
    End If
End If
End Function


Private Sub CalculateHeuristics(ByVal StartX As Long, StartY As Long, ByVal EndX As Long, ByVal EndY As Long)
                                
Dim X As Long

' this is the A* portion of the pathfinding routine
' By calculating which rectangles are relatively closer to the distant end object,
' we prioritize the closest when we choose which node to expand next

' All other pathfiding routines rely on a grid of equally sized cells,
' well window regions obviously aren't built that way. So we use our own algorithm

' Probably the most important calculation routine in this project.
' The better the heuristics, the more likely a better path is chosen first;
' which in turn can abort longer paths before they spawn too many child legs

For X = 0 To UBound(Legs)

    ' First calculate the vertical distance to the top or bottom edge of the rectangle
    ' whichever is closer to the distant end
    
    If NodeRects(X).Top < EndY Then
        Legs(X).Heuristics(0) = Abs(EndY - NodeRects(X).Bottom)
    Else
        Legs(X).Heuristics(0) = Abs(EndY - NodeRects(X).Top)
    End If
    If NodeRects(X).Top < StartY Then
        Legs(X).Heuristics(1) = Abs(StartY - NodeRects(X).Bottom)
    Else
        Legs(X).Heuristics(1) = Abs(StartY - NodeRects(X).Top)
    End If
    
    ' next determine if the rectangle can be seen directly above/below the distant end
    ' and if so, then add nothing more to the heuristic; otherwise add the horizontal distance
    ' from the closest edge (right or left) to the distant end
    If PtInRect(NodeRects(X), EndX, NodeRects(X).Top) = 0 Then
    
        If Abs(NodeRects(X).Left - EndX) < Abs(NodeRects(X).Right - EndX) Then
            Legs(X).Heuristics(0) = Legs(X).Heuristics(0) + Abs(NodeRects(X).Left - EndX)
        Else
            Legs(X).Heuristics(0) = Legs(X).Heuristics(0) + Abs(NodeRects(X).Right - EndX)
        End If
    
    End If
    If PtInRect(NodeRects(X), StartX, NodeRects(X).Top) = 0 Then
    
        If Abs(NodeRects(X).Left - StartX) < Abs(NodeRects(X).Right - StartX) Then
            Legs(X).Heuristics(1) = Legs(X).Heuristics(1) + Abs(NodeRects(X).Left - StartX)
        Else
            Legs(X).Heuristics(1) = Legs(X).Heuristics(1) + Abs(NodeRects(X).Right - StartX)
        End If
    
    End If
    ' reset the owner & explored values
    Legs(X).Owner = 0
    Legs(X).CollideIndex = 0
Next
End Sub

Private Function ExtractRectangles(rgnMap As Long) As Long

' Homegrown function to extract rectangle structure of a windows region

Dim rSize As Long, vRgnData() As Byte, X As Long

    ' 1st get the buffer size needed to return rectangles info from this region
    rSize = GetRegionData(rgnMap, ByVal 0&, ByVal 0&)
    If rSize > 0 Then   ' success
        ' create the buffer & call function again to fill the buffer
        ReDim vRgnData(0 To rSize - 1) As Byte
        If rSize = GetRegionData(rgnMap, rSize, vRgnData(0)) Then     ' success
        
            ' Here are some tips for the structure returned
            ' Bytes 8-11 are the number of rectangles in the region
            ' Bytes 12-15 is structure size information -- not important for what we need
            ' Bytes 16-31 are the bounding rectangle's dimensions
            ' Bytes 32 to end of structure are the individual rectangle's dimensions
            ' The rectangle structure (RECT) is 16 bytes or LenB(RECT)
        
            ' Let's retrieve the number of rectangles in the structure (b:8-11)
            CopyMemory rSize, vRgnData(8), ByVal 4&
            ReDim NodeRects(0 To rSize - 1)
            CopyMemory NodeRects(0), vRgnData(32), rSize * 16
        End If
    End If

ExtractRectangles = rSize
Erase vRgnData
End Function

Private Sub DisplayPartialPath(ByVal StartIndex As Long)
' used for debugging purposes only, can be removed/rem'd out
If StartIndex > UBound(vPaths) Or StartIndex < 1 Then Exit Sub
Dim lColor As Long, X As Integer, Y As Integer, X1 As Integer, Y1 As Integer
With frmMain
    lColor = .picWork.ForeColor
    .picWork.ForeColor = vbYellow
    X = vPaths(StartIndex).X
    Y = vPaths(StartIndex).Y
    Do
        StartIndex = vPaths(StartIndex).Parent
        X1 = vPaths(StartIndex).X
        Y1 = vPaths(StartIndex).Y
        frmMain.picWork.Line (X, Y)-(X1, Y1)
        'Debug.Print X; Y, X1; Y1
        X = X1
        Y = Y1
    Loop Until vPaths(StartIndex).Parent < 0
    .picWork.ForeColor = lColor
    .picWork.Refresh
End With
End Sub

Private Sub DisplayRectangle(Index As Long)

' used for debugging purposes only, can be removed/rem'd out
Dim lColor As Long
If Index > UBound(NodeRects) Or Index < 0 Then Exit Sub
With NodeRects(Index)
    lColor = frmMain.picWork.ForeColor
    frmMain.picWork.ForeColor = vbGreen
    Rectangle frmMain.picWork.hdc, .Left, .Top, .Right, .Bottom
End With
frmMain.picWork.ForeColor = lColor
End Sub

Private Sub BuildReturnArray(inArray() As Long)
' this function sets the A & Z side path segments in order

Dim pIndex As Long, arrayCt As Long
' first we'll count how many segments are in the A-Side path
pIndex = FullPath.aIndex
Do
    arrayCt = arrayCt + 1
    pIndex = vPaths(pIndex).Parent
Loop Until pIndex < 0

' now we'll dimension our return array & copy the
' XY coords from the A-side path into the array
pIndex = FullPath.aIndex
arrayCt = arrayCt - 1
ReDim Preserve inArray(0 To arrayCt)
arrayCt = 0
Do
   CopyMemory inArray(UBound(inArray) - arrayCt), MakeLong(vPaths(pIndex).X, vPaths(pIndex).Y), 4&
   pIndex = vPaths(pIndex).Parent
   arrayCt = arrayCt + 1
Loop Until pIndex < 0

' next we simply copy the Z-side paths into the array
pIndex = FullPath.zIndex
Do
    ReDim Preserve inArray(0 To arrayCt)
    CopyMemory inArray(arrayCt), MakeLong(vPaths(pIndex).X, vPaths(pIndex).Y), 4&
    arrayCt = arrayCt + 1
    pIndex = vPaths(pIndex).Parent
Loop Until pIndex < 0
End Sub

Private Sub AnimatePts()
' used for testing purposes only, can be removed/rem'd out

Dim X As Long, Y As Long, AniIndex As Integer, aniRect As RECT
X = vPaths(UBound(vPaths)).X
Y = vPaths(UBound(vPaths)).Y
If ((vPaths(UBound(vPaths)).Status And 1) = 1) Then AniIndex = 1
With frmMain.cmdAni(AniIndex)
    X = X - (.Width \ 2)
    Y = Y - (.Height \ 2)
    SetRect aniRect, .Left, .Top, .Left + .Width, .Top + .Height
    .Move X, Y
    RedrawWindow frmMain.picMap.hwnd, aniRect, 0&, True
End With
Sleep 10
End Sub
