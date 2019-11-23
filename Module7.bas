Option Explicit

Public strText_start As Date
Public strText_end As Date
Public strPath As String

Sub SaveAttachmentFile()

    Dim myNamespace As NameSpace
    Dim myInbox As Object
    Dim mySubfolder As Object
    'Dim target As String
    Dim arr() As String
    Dim target As Variant '配列の結果格納用
    Dim i As Integer 'For-Next用
    Dim r As Integer 'Ascでreplaceするときのカウント用
    
    'Application.Run ("'UserForm1'!CommandButton1_Click")
    
    UserForm1.Show
    Debug.Print strText_start
    Debug.Print strText_end


    
    '以下、サブフォルダの指定
    Set myNamespace = GetNamespace("MAPI")
    Set myInbox = myNamespace.GetDefaultFolder(olFolderInbox)
    Set mySubfolder = myInbox.Folders.Item("test")

    'myInbox.Items.Add ("IPM.Note.Customer")
     'UserPropertyオブジェクトのValueプロパティでユーザー定義フィールドの値を設定
      'myInbox.UserProperties("差出人").Value = "test"
      'myInbox.UserProperties("住所1").Value = "沖縄県那覇市首里金城町1-2-3"
      'myInbox.UserProperties("住所2").Value = "テストビル10階"

    target = mySubfolder.Items(1).ReceivedTime 'サブフォルダ内の日付取得
    Erase arr
    ReDim arr(2)
    arr = Split(target) '文字列を区切り文字で分割する
    
    
    
    'Debug.Print arr(0)
    'Debug.Print "" & arr(0) & ""
    
    'Call CommandButton1_Click
    'For i = LBound(arr) To UBound(arr)
                   
    'Next i
    
    'Debug.Print arr(0)
    Dim nowitem As MailItem 'メールアイテム
    Dim nowdate As String
    Dim nowsbjct As String
    Dim flname As String
    
'ここからコメント↓
'    For i = 0 To mySubfolder.Items.Count
'        If "" & arr(0) & "" >= strText_start And "" & arr(0) & "" <= strText_end Then
'            Set nowitem = mySubfolder.Items(i + 1) 'i番目のアイテムの定義
'
'            nowsbjct = nowitem.Subject '件名取得
'
            Dim ary As Variant '不可文字削除用配列
'
            ary = Array("RE: ", "FW: ", "\", "/", ":", "*", "?", "<", ">", "|", """", "[", "]", ",")
            
''            nowsbjct = nowitem.Subject '件名
'
'            For r = 0 To UBound(ary) 'UBoundで配列の最大の要素数
'                        nowsbjct = Replace(nowsbjct, ary(r), Asc(ary(r))) 'replace(置換の対象、検索、置き換える文字)
'            Next r
'
'            flname = nowsbjct & ".msg" 'ファイル名生成
'
'            nowitem.SaveAs strPath & "\" & flname, olMSG
'ここまでコメント↑

'以下、FileSystemObjectの例だけど関係ない
'            Dim fso As FileSystemObject
'            Set fso = New FileSystemObject
'
'            Const Path As String = "C:\Users\Ryosuke\Desktop\test\"
'
'            Dim f As File
'            Set f = fso.GetFile(Path & "*.*")
'
'            Dim s As String
'            s = f.Name  'ファイル名の取得
'
'            Do While f <> ""
'                Cnt = Cnt + 1
'                For r = 0 To UBound(ary) 'UBoundで配列の最大の要素数
'                    buf = Replace(buf, ary(r), Chr(ary(a)))     'replace(置換の対象、検索、置き換える文字)
'                Next r
'            Loop

            

            Dim buf As String, Cnt As Long

            Const Path As String = "C:\Users\Ryosuke\Desktop\test\"
            
            buf = Dir(Path & "*.msg")
            
            Do While buf <> ""
'                Cnt = Cnt + 1
'
                For r = 0 To UBound(ary) 'UBoundで配列の最大の要素数
                    Dim s As String 'テスト用にとりあえず用意したやつ
                    s = Asc(ary(r)) 'テスト用に用意したやつ、格納時の状態

                    buf = Replace(buf, s, Chr(s))     'replace(置換の対象、検索、置き換える文字)
                Next r
                Debug.Print buf
                
                buf = Dir()
                
         Loop


            
            
            
            'ファイル名で格納するさいに使えないものがあるので、そこのリプレースする処理を書く
            'i番目のアイテムの定義のところがアイテムの最後までいくとエラーになる。考える
            
            'mySubfolder.Items(i + 1).SaveAs strPath
            'Debug.Print arr(0)
             
        'End If 対応するため現在コメント
    'Next i 対応するため現在コメント
    
    'Set fso = Nothing
    Set nowitem = Nothing

End Sub
