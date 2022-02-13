'#######################################################################################
'# <説明>
'# 指定のIT機器に対してPingコマンドを使用した死活監視を行う
'# 
'# <更新日>
'# 作成日：20210814
'# 最終更新日：yyyymmdd
'#
'# <使用時における注意事項>
'# ・文字コードは「SJIS(ANSI)」を指定すること
'#
'# <コメント>
'# ・"<<>>"で囲っている部分は修正箇所となっていますので、ご自身の環境に合わせて記載してください
'#
'#######################################################################################


'################################################################################
'## 事前準備
'################################################################################

'----- 宣言 -----
Dim strDate
Dim strDate_2

'----- 日付(yyyymmdd - hhmm)取得  -----
strDate_2 = Now()
strDate_2 = Left(strDate_2, 16)
strDate_2 = Replace(strDate_2, "/", "")

'----- 日付(yyyymmdd)取得  -----
strDate = Date()
strDate = Replace(strDate, "/", "")


'################################################################################
'## パラメーター指定
'################################################################################

'----- 宣言 -----
Dim target_host
Dim objWMIService
Dim PingSet
Dim Ping

'----- ターゲットのホスト名を指定　※監視ホストのホスト名を記載する -----
target_host = "<<ホスト名を指定>>"

'----- 実行ログ -----
'// 実行ログ出力先フォルダ
Dim fileNameCrown :fileNameCrown = "<<ログファイル格納先ディレクトリを指定>>"
'// 実行ログ出力先ファイル名
Dim fileNameTail :fileNameTail = "_ping_healthcheck"
'// 実行ログ出力先ファイル拡張子
Dim fileNameSuffix :fileNameSuffix = ".log"
'// ファイル名連結
fileName = fileNameCrown & strDate & fileNameTail & fileNameSuffix

'----- 監視ターゲットを指定する　※3行目末尾に監視ホストのIPを指定する -----
Set objWMIService = GetObject("winmgmts:\\.")
Set PingSet = objWMIService.ExecQuery _
("Select * From Win32_PingStatus Where Address = '<<xx.xx.xx.xx>>'")


'################################################################################
'## 指定したIPに対してPingを実行
'################################################################################

'----- 上記で指定したターゲットに対してPingコマンドを実行する -----
For Each Ping In PingSet
  Select Case Ping.StatusCode
  Case 0
    checkPing = True
  Case 11010
    checkPing = False
  End Select
Next


'################################################################################
'## 実行ログの出力先ファイルを作成
'################################################################################

'----- 宣言 -----
Dim FSO
Dim oLog

'----- オブジェクト生成 -----
Set FSO = CreateObject("Scripting.FileSystemObject")

'----- リソース作成処理作成 -----
If FSO.FileExists(fileName) Then
  '// 存在すれば何もしない
Else
  '// ファイル存在しなければ作成 -----
  Set oLog = FSO.CreateTextFile(fileName)
End If

'----- ファイル操作後処理 -----
Set oLog = Nothing
Set FSO = Nothing


'################################################################################
'## メイン処理（ログ出力、メール通知）
'################################################################################

'----- pingの実行結果により処理を変更する -----
If checkPing = True Then
  Success_Log_Output
Else
  Failed_Log_Output
  Send_Mail
End if


'################################################################################
'## 関数：ログ出力
'################################################################################

'----- Ping成功時のログ出力処理 -----
function Success_Log_Output() 
  '// エラー処理
  On Error Resume Next
  '// 宣言
  Dim objFSOO
  Dim objFile
  '// オブジェクト生成
  Set objFSOO = WScript.CreateObject("Scripting.FileSystemObject")
  '// pingの実行結果をログに出力する
  If Err.Number = 0 Then
    Set objFile = objFSOO.OpenTextFile(filename, 8, True)
      If Err.Number = 0 Then
        objFile.WriteLine(strDate_2 & "　Ping was successful")
        objFile.Close
      Else
        WScript.Echo "ファイルオープンエラー: " & Err.Description
      End If
  Else
    WScript.Echo "エラー: " & Err.Description
  End If
    Set objFile = Nothing
    Set objFSOO = Nothing
Else
end function
  
'----- Ping失敗時のログ出力処理 -----
function Failed_Log_Output()
 '// エラー処理
  On Error Resume Next
  '// 宣言
  Dim objFSOO_Fai
  Dim objFile_Fai
  Dim Failed_messa
  '// オブジェクト生成
  Set objFSOO_fai = WScript.CreateObject("Scripting.FileSystemObject")
  '// Failedメッセージ
  Failed_messa = "　Ping was Failed The health check of the target terminal failed. Investigate and identify the cause. Also, an email was sent to the specified address for Failed Message. If you have any questions about how to deal with it, contact the server administrator. Please go"
  '// pingの実行結果をログに出力する
  If Err.Number = 0 Then
    Set objFile_Fai = objFSOO_Fai.OpenTextFile(filename, 8, True)
    If Err.Number = 0 Then
      objFile_Fai.WriteLine(strDate_2 & Failed_messa)
      objFile_Fai.Close
    Else
      WScript.Echo "ファイルオープンエラー: " & Err.Description
    End If
  Else
    WScript.Echo "エラー: " & Err.Description
  End If
  '// オブジェクト解放
  Set objFile = Nothing
  Set objFSOO = Nothing
end function
 

'################################################################################
'## 関数：メール通知
'################################################################################

function Send_Mail()
  '----- 宣言 -----
  Dim mailbody

  '----- オブジェクト生成 -----
  Set objMail = CreateObject("CDO.Message")

  '----- メールの本文を生成  -----
  mailBody = "[Critical] Failed to monitor the life and death of the target host" & Target_HOST & vbCrLf & vbCrLf
  mailBody = mailBody & "Failed to monitor the life and death of the target host" & vbCrLf
  mailBody = mailBody & "Access the target host and check if the host has started normally" & vbCrLf
  mailBody = mailBody & "Please refer to the incident management table for how to deal with it" & vbCrLf & vbCrLf

  '----- 情報取得　※通知を受け取るメールアドレスを記載する -----
  objMail.From = "****.****@gmail.com"
  objMail.To = "****.****@gmail.com"
  objMail.Subject = "Failed to monitor the life and death of the target host" & target_host
  objMail.TextBody = mailbody

  '----- メール作成  ※Googleアカウント情報を記載する -----
  strConfigurationField ="http:-----schemas.microsoft.com/cdo/configuration/"
  With objMail.Configuration.Fields
    .Item(strConfigurationField & "sendusing") = 2
    .Item(strConfigurationField & "smtpserver") = "smtp.googlemail.com"
    .Item(strConfigurationField & "smtpserverport") = 465
    .Item(strConfigurationField & "smtpusessl") = true
    .Item(strConfigurationField & "smtpauthenticate") = 1
    .Item(strConfigurationField & "sendusername") = "******.******@gmail.com"
    .Item(strConfigurationField & "sendpassword") = "***************"
    .Item(strConfigurationField & "smtpconnectiontimeout") = 60
    .Update
  end With
  '----- メール送信 -----
  objMail.Send

  '----- オブジェクト開放 -----
  Set objMail = Nothing

end function