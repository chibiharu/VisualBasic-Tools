Option Explicit

'--- 宣言 ---
  Dim objIE
  Dim el
  Dim objbutton

'--- オブジェクト生成 ---
  Set objIE = CreateObject("InternetExplorer.Application")
  objIE.Visible = True

'--- IEを開く ---
'※ [example.com]をログインしたいサイトURLに変えてください
  objIE.navigate "example.com"

'--- ページが読み込まれるまで待つ ---
  Do While objIE.Busy = True Or objIE.readyState <> 4
    WScript.Sleep 100        
  Loop

'--- a要素にログインがあったらクリックする ---
'※ 環境に合わせて【ログイン】を変えて下さい
  For each el In objIE.document.Links             
    if instr(el.innerText,"���O�C��") then
      el.click
      exit for    
    end if
  next    

'--- ページが読み込まれるまで待つ ---
  Do While objIE.Busy = True Or objIE.readyState <> 4
    WScript.Sleep 100        
  Loop  

'--- IDとパスワードを入力する ---
'※ Elementを環境に合わせて変えて下さい
'※ valueをご自身のIDとpasswordに変えて下さい
  With objIE.document       
    .getElementsByName("login_element")(0).Value = "login_id" 
    .getElementsByName("password_element")(0).Value = "password"        
  End With     

'--- button要素をコレクションとして取得 ---
'※ Elementを環境に合わせて変えて下さい
  Set objbutton = objIE.document.getElementbyid("button_id")
  
'--- 「ログイン」ボタンクリック ---
  objbutton.click