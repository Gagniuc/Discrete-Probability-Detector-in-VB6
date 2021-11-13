VERSION 5.00
Begin VB.Form DPD 
   Caption         =   "Discrete Probability Detector"
   ClientHeight    =   10830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   LinkTopic       =   "Form1"
   ScaleHeight     =   722
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   840
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox InP 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Text            =   "c13e13da2073260c2194c15d782e86a9"
      Top             =   240
      Width           =   8535
   End
   Begin VB.TextBox OutPut 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "DPD.frx":0000
      Top             =   720
      Width           =   12375
   End
   Begin VB.CommandButton MakeMatrix 
      Caption         =   "Extract"
      Height          =   375
      Left            =   9480
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label InputLabel 
      Caption         =   "Input:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "DPD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###################################################################################################
'# John Wiley & Sons, Inc.                                                                         #
'#                                                                                                 #
'# Book:   Markov Chains: From Theory To Implementation And Experimentation                        #
'# Author: Dr. Paul Gagniuc                                                                        #
'# Data:   01/05/2017                                                                              #
'#                                                                                                 #
'# Title:                                                                                          #
'# Discrete Probability Detector                                                                   #
'#                                                                                                 #
'# Short Description:                                                                              #
'# The purpose of this algorithm is to convert any string into a probability matrix.               #
'#_______________________                                                                          #
'# Detailed description: \                                                                         #
'# This algorithm is an advanced variation of the "ExtractProb" function from the book. The        #
'# main difference between "ExtractProb" function and the DPD algorithm is the automatic           #
'# identification of states.                                                                       #
'#_________________________________________________________________________________________________#
'# Initially, the states are identified in the first phase. Each new letter found in "S" is        #
'# appended to the string forming in variable "a". Thus, variable "a" gradually increases until    #
'# all types of letters from "S" are identified.                                                   #
'#_________________________________________________________________________________________________#
'# In the second phase the elements of matrix "m" are filled with zero values for later            #
'# purposes. Also, in the second phase the first column of matrix "e" is filled with letters       #
'# found in variable "a", and the second column of matrix "e" is filled with zero values for       #
'# later use.                                                                                      #
'#_________________________________________________________________________________________________#
'# In the third phase, the transitions between letters of "S" are counted and stored in            #
'# matrix "m".                                                                                     #
'# The strategy in this particular case is to fill matrix "m" with transition counts before the    #
'# last letter in "S" is reached. In this case, the first column of matrix "e" already contains    #
'# the letters from variable "a". The two components of vector "l" contain the "i" and "i+1"       #
'# letters from "S". The count of individual transitions between letters is made by a comparison   #
'# between vector "l" and the elements from the first column of matrix "e". The number of rows     #
'# in matrix "m" and matrix "e" is the same, namely "d". Therefore, an extra loop can be avoided   #
'# by mapping matrix "m" through a coordinate system. For instance, if the letter from position    #
'# "i" in "S" stored in "l(0)" and the letter from "j" row in matrix "e" (e(j,0)) are the same     #
'# then variable "r = j". Likewise, if the letter "i+1" stored in l(1) and the letter from e(j,0)  #
'# are the same then variable "c = j". Variable "r" represents the rows of matrix "m", whereas     #
'# variable "c" represents the columns of matrix "m" (m(r, c)).                                    #
'# Thus, at each step through "S", an element of matrix "m" is always incremented according to     #
'# the coordinates received from "r" and "c". This "coordinate" approach greatly increases the     #
'# processing speed of the algorithm. The number of loops = (k-1)*d, where "d" represents the      #
'# number of states (or letter types), and "k" is the number of letters in "S". When the letter    #
'# stored in "l(0)" and the letter from "j" row in matrix "e" are the same, the second column of   #
'# matrix "e" is also incremented. The second column of matrix "e" stores the number of            #
'# appearances for each type of letter in "S".                                                     #
'#_________________________________________________________________________________________________#
'# In the fourth phase, the counts from matrix "m" elements are divided by the counts from the     #
'# second column of matrix "e". The result of this division is stored in the same position in      #
'# matrix "m", and represents a transition probability.                                            #
'#_________________________________________________________________________________________________#
'#                                                                                                 #
'# Special considerations:                                                                         #
'#                                                                                                 #
'# If a state at the end of "S" (ie HAHAAAHQ) does not occur in the rest of "S" then matrix "m"    #
'# will contain a row with all elements on zero. Since it is at the end of "S", the letter does    #
'# not make a transition to anything. If a state from the beginning of "S" (ie. QHAHAAAH) does     #
'# not occur in the rest of "S" then matrix "m" will contain a column with all elements on zero.   #
'# Since the first letter it is only seen at the beginning of "S", no other letter makes a         #
'# transition to it.                                                                               #
'#                                                                                                 #
'# The meaning of variables:                                                                       #
'#      _____________________________________________________________________________________      #
'# S = |The string that is being analyzed.                                                  _|     #
'#      _____________________________________________________________________________________      #
'# q = |It is a flag variable with initial value of 1. The value of q becomes zero only if a |     #
'#     |letter x in the "S" string coresponds with a letter y in the "a" string.            _|     #
'#      _____________________________________________________________________________________      #
'# a = |The variable that holds the letters representing the states. The variable gradually  |     #
'#     |increases in length as the "S" string is analyzed. At each loop, a new letter is     |     #
'#     |added to variable "a" only if the value of q becomes zero. Thus, at the end of the   |     #
'#     |first loop the number of letters in the variable is equal to the total number        |     #
'#     |of states.                                                                          _|     #
'#      _____________________________________________________________________________________      #
'# d = |Represents the total number of states and is the length of "a" variable.            _|     #
'#      _____________________________________________________________________________________      #
'# m = |The main probability matrix which the function produces.                            _|     #
'#      _____________________________________________________________________________________      #
'# k = |Represents the length of the "S" string.                                            _|     #
'#      _____________________________________________________________________________________      #
'# e = |It is a matrix with two columns, namely column 0 and 1. Column 0 stores all the      |     #
'#     |letters found in "a". Column 1 stores the number of appearances for each type        |     #
'#     |of letter in "S".                                                                   _|     #
'#      _____________________________________________________________________________________      #
'# l = |Is a vector with two components. Vector "l" contains the "i" and "i+1" letters       |     #
'#     |from "S".                                                                           _|     #
'#                                                                                                 #
'###################################################################################################

Private Sub MakeMatrix_Click()
    Discrete_Probability_Detector (InP.Text)
End Sub

Function Discrete_Probability_Detector(ByVal S As String)

    Dim e() As String
    Dim m() As String
    Dim l(0 To 1) As String
    
    k = Len(S)
    w = 1
    
    For i = 1 To k
        q = 1
        For j = 0 To Len(a)
            x = Mid(S, i, 1)
            y = Mid(a, j + 1, 1)
            If x = y Then q = 0
        Next j
        If q = 1 Then a = a & x
    Next i

    d = Len(a)

    ReDim e(w To d, 0 To 1) As String
    ReDim m(w To d, w To d) As String
    
    For i = w To d
        For j = w To d
          m(i, j) = 0
          If j = w Then
            e(i, 0) = Mid(a, i, 1)
            e(i, 1) = 0
          End If
        Next j
    Next i

    For i = 1 To k - 1
        l(0) = Mid(S, i, 1)
        l(1) = Mid(S, i + 1, 1)
        For j = w To d
            If l(0) = e(j, 0) Then
               e(j, 1) = Val(e(j, 1)) + 1
               r = j
            End If
            If l(1) = e(j, 0) Then c = j
        Next j
        m(r, c) = Val(m(r, c)) + 1
    Next i
    
    tmp = "S=" & S & vbCrLf & vbCrLf
    tmp = tmp & "The algorithm detected a total of " & (d - w + 1) & " states." & vbCrLf & vbCrLf
    tmp = tmp & MatrixPaint(w, d, m, a, "(C)", "Count:") & vbCrLf
    
    For i = w To d
        For j = w To d
            If Val(e(i, 1)) > 0 Then
            m(i, j) = Val(m(i, j)) / Val(e(i, 1))
            End If
        Next j
    Next i

    tmp = tmp & MatrixPaint(w, d, m, a, "(P)", "Transition matrix M:")

    OutPut.Text = tmp

End Function


Function MatrixPaint(w, d, ByVal m As Variant, a, n, ByVal msg As String) As String
    
    Dim e() As String
    ReDim e(1 To d) As String
    
    d = Len(a)
    q = "|     "
    h = "|_____|"
    l = vbCrLf
    
    For i = (w - 1) To d
        If i = (w - 1) Then t = t & l & "."
        t = t & "_____."
        If i = d Then t = t & l & "| " & n & " |  "
    Next i

    For i = w To d
        e(i) = Mid(a, i, 1)
        t = t & e(i) & "  |  "
        h = h & "_____|"
    Next i
    
    t = t & l & h & l
    
    For i = w To d
        For j = w To d
            v = Round(m(i, j), 2)
            u = Mid(q, 1, Len(q) - Len(v))
            If j = d Then o = "|" Else o = ""
            For b = w To d
                If j = w And i = b Then
                    t = t & "|  " & e(i) & "  "
                End If
            Next b
            t = t & u & v & o
        Next j
    t = t & l & h & l
    Next i
    
    MatrixPaint = msg & " M[" & Val(d - w + 1) & "," & Val(d - w + 1) & "]" & l & t & l
    
End Function


Private Sub Form_Resize()
    If DPD.ScaleWidth > 0 Then
        OutPut.Width = DPD.ScaleWidth - OutPut.Left - 10
        OutPut.Height = DPD.ScaleHeight - OutPut.Top - 10
    End If
End Sub
