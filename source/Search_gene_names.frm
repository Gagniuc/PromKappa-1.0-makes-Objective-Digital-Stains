VERSION 5.00
Begin VB.Form Search_gene_names 
   Caption         =   "Search"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form2"
   ScaleHeight     =   6225
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Gene name lenght"
      Height          =   2055
      Left            =   2640
      TabIndex        =   4
      Top             =   720
      Width           =   2895
      Begin VB.TextBox max_len 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Text            =   "10"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox min_len 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Text            =   "2"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Max lenght"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Min lenght"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CheckBox JV 
      Caption         =   "Just verify the list"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton ok 
      Caption         =   "Ok"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox list_gene 
      Height          =   4575
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Paste the gene names:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Search_gene_names"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'  ________________________________                          _____________________
' /  PromKappa                     \________________________/       v2.00         |
' |                                                                               |
' |            Name:  PromKappa                                                   |
' |        Category:  Open source software                                        |
' |          Author:  Paul A. Gagniuc                                             |
' |           Email:  paul_gagniuc@acad.ro                                        |
' |                                                                               |
' |    Date Created:  3/01/2012                                                   |
' |       Tested On:  WinXP, WinVista, Win7, Win8                                 |
' |             Use:  Analysis of gene promoters                                  |
' |                                                                               |
' |                  _____________________________                                |
' |_________________/                             \_______________________________|


Private Sub ok_Click()
Dim lg() As String

pro_no = PromKappa.List_promotori.ListCount
lg = Split(list_gene, vbCrLf)

Item_found = 0

For genes = 0 To UBound(lg)

PromKappa.gene_search_txt.Text = lg(genes)

If Len(lg(genes)) >= Val(min_len.Text) And Len(lg(genes)) <= Val(max_len.Text) Then

For i = 0 To pro_no

    If PromKappa.List_promotori.List(i) <> "" Then

        If InStr(UCase(PromKappa.List_promotori.List(i)), UCase(PromKappa.gene_search_txt.Text)) <> 0 Then
        
            a = PromKappa.List_promotori.ItemData(i)
            Item_found = Item_found + 1
            
            If JV.Value = 0 Then
                PromKappa.List_promotori.Enabled = False

                PromKappa.secventata.Text = Split(PromKappa.List_promotori.List(i), "|")(1)
                PromKappa.seq_name_from_file.Caption = Split(PromKappa.List_promotori.List(i), "|")(0)
                PromKappa.call_sub_Procesare_Click

                PromKappa.List_promotori.Enabled = True
        
            End If
        
        Exit For
        End If
        

    End If


Next i

End If

Next genes
MsgBox "Total genes in list: " & UBound(lg) & vbCrLf & "Total promoter found: " & Item_found & " genes"
End Sub
