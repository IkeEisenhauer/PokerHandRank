Function PokerHand(Card1,Card2,Card3,Card4,Card5)

Dim myHand(5)
Dim mySuitLessHand(5)
Dim mySameKind
Dim iTemp

myHand(1) = Card1
myHand(2) = Card2
myHand(3) = Card3
myHand(4) = Card4
myHand(5) = Card5

HandCode = "0"
RankCode = "xxxxx"

For i = 1 To 5
    For j = i + 1 To 5
        If myHand(i) > myHand(j) Then
            iTemp = myHand(j)
            myHand(j) = myHand(i)
            myHand(i) = iTemp
        End If
    Next
Next 

For i = 1 To 5
mySuitLessHand(i) = ((myHand(i) - 1) Mod 13) + 1
If mySuitLessHand(i) = 1 Then
    mySuitLessHand(i) = 14
End If
Next

For i = 1 To 5
    For j = i + 1 To 5
        If mySuitLessHand(i) > mySuitLessHand(j) Then
            iTemp = mySuitLessHand(j)
            mySuitLessHand(j) = mySuitLessHand(i)
            mySuitLessHand(i) = iTemp
        End If
    Next
Next

mySameKind = SameKind(mySuitLessHand)

If IsFlush(myHand) Then
    If IsStraight(mySuitLessHand) Then
        HandCode = "8"
    Else
        HandCode = "5"
    End If
ElseIf IsStraight(mySuitLessHand) Then
        HandCode = "4"
ElseIf mySameKind = 6 Then
        HandCode = "7"
        If mySuitLessHand(5) <> mySuitLessHand(4) Then
            mySuitLessHand(1) = mySuitLessHand(5)
            mySuitLessHand(5) = mySuitLessHand(4)
        End If
ElseIf mySameKind = 4 Then
        HandCode = "6"
        If mySuitLessHand(5) <> mySuitLessHand(3) Then
            mySuitLessHand(1) = mySuitLessHand(5)
            mySuitLessHand(2) = mySuitLessHand(5)
            mySuitLessHand(4) = mySuitLessHand(3)
            mySuitLessHand(5) = mySuitLessHand(3)
        End If
ElseIf mySameKind = "3" Then
        HandCode = "3"
        If mySuitLessHand(2) = mySuitLessHand(3) Then
            mySuitLessHand(2) = mySuitLessHand(5)
            mySuitLessHand(5) = mySuitLessHand(3)
            If mySuitLessHand(1) = mySuitLessHand(3) Then
                mySuitLessHand(1) = mySuitLessHand(4)
                mySuitLessHand(4) = mySuitLessHand(3)
            End If
        End If
ElseIf mySameKind = "2" Then
        HandCode = "2"
        If mySuitLessHand(1) = mySuitLessHand(2) Then
            If mySuitLessHand(3) = mySuitLessHand(4) Then
                mySuitLessHand(3) = mySuitLessHand(5)
                mySuitLessHand(5) = mySuitLessHand(4)
            End If
            mySuitLessHand(1) = mySuitLessHand(3)
            mySuitLessHand(3) = mySuitLessHand(2)
        End If
ElseIf mySameKind = "1" Then
        HandCode = "1"
        If mySuitLessHand(1) = mySuitLessHand(2) Then
            mySuitLessHand(1) = mySuitLessHand(3)
            mySuitLessHand(3) = mySuitLessHand(2)
        End If
        If mySuitLessHand(2) = mySuitLessHand(3) Then
            mySuitLessHand(2) = mySuitLessHand(4)
            mySuitLessHand(4) = mySuitLessHand(3)
        End If
        If mySuitLessHand(3) = mySuitLessHand(4) Then
            mySuitLessHand(3) = mySuitLessHand(5)
            mySuitLessHand(5) = mySuitLessHand(4)
        End If
Else
    HandCode = "0"
End If

RankCode = RankCoder(mySuitLessHand)

PokerHand= HandCode & RankCode

End Function

Function RankCoder(ByVal Hand) As String

RankCoder = ""

For j = 5 To 1 Step -1
    RankCoder = RankCoder & Mid("XABCDEFGHIJKLM", Hand(j), 1)
Next

End Function
Function IsFlush(ByVal Hand) As Integer

IsFlush = 0

If (((Hand(5) - Hand(1)) < 13) And ((Hand(5) Mod 13) > (Hand(1) Mod 13))) Then IsFlush = True

End Function
Function IsStraight(ByVal Hand) As Integer

IsStraight = False
    
If (Hand(1) + 1 = Hand(2)) Then
    If Hand(2) + 1 = Hand(3) Then
        If Hand(3) + 1 = Hand(4) Then
            If (Hand(4) + 1 = Hand(5)) Or (Hand(4) = "5" And Hand(5) = "14") Then
                IsStraight = True
                Exit Function
            End If
        End If
      End If
End If

End Function
Function SameKind(ByVal Hand)
   
For i = 1 To 4
    For j = i + 1 To 5
        If Hand(i) = Hand(j) Then
            SameKind = SameKind + 1
        End If
    Next
Next 

End Function
