'***********************************************************
' * All Rights Reserved, Copyright(C) korai 2015
' * プログラム名：  動画管理システム
' * プログラムＩＤ：common.vbs
' * 作成日付：	   2015/04/30
' * 最終更新日：	   2015/04/30
' * プログラム説明：共通モジュール
' * 変更履歴
' * 変更日時		変更者名			説明
' * 2015/04/30	紅雷				新規作成
'************************************************************

'***********************************************************
'初期画面のhtmlを出力する
'[IN]	なし
'[OUT]	MovieListのIDが振られた個所にstrStartHtmlの内容を書き込む
'***********************************************************
	Sub startDisplay()
		Dim strStartHtml
		strStartHtml ="<table width='85%'><tr><td class='top'>TOPメニュー</td>" & _
		"<td align='right'><input class='search' type='text' id='searchText'>" & _
		"<input class ='searchbtn' type='button' value='検索' onClick='ReadSearchFiles()'>" & _
		"</td></tr></table><p>"

		strStartHtml = strStartHtml & "<table><tr>" & _
		"<td><input  type ='button' value='あ' onClick='ReadMovieFolder " & chr(34) & "あ" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  class ='tile' type ='button' value='か'  onClick='ReadMovieFolder " & chr(34) & "か" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  type ='button' value='さ'  onClick='ReadMovieFolder " & chr(34) & "さ" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  class ='tile' type ='button' value='た'  onClick='ReadMovieFolder " & chr(34) & "た" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  type ='button' value='な'  onClick='ReadMovieFolder " & chr(34) & "な" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"</tr>" & _
		"<tr>" & _
		"<td><input  class ='tile' type ='button' value='は'  onClick='ReadMovieFolder " & chr(34) & "は" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  type ='button' value='ま'  onClick='ReadMovieFolder " & chr(34) & "ま" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  class ='tile' type ='button' value='や'  onClick='ReadMovieFolder " & chr(34) & "や" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  type ='button' value='ら'  onClick='ReadMovieFolder " & chr(34) & "ら" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  class ='tile' type ='button' value='わ'  onClick='ReadMovieFolder " & chr(34) & "わ" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"</tr>" & _
		"<tr>" & _
		"<td><input  type ='button' value='その他'  onClick='ReadMovieFolder " & chr(34) & "その他" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"</tr></table>" & _
		"<p><table  width=750px ><tr><td align='right'><input class ='nomal' type ='button' value='終了' onClick='window.close()'></td></tr></table>"

		call readHtml("MovieList",strStartHtml)
	End Sub
'***********************************************************
'フォルダ読み込んで以下に存在するファイル、もしくはフォルダからhtmlを出力する
'[IN]	obj　該当フォルダ名
'[IN]	flg　0：ファイル　1：フォルダ
'[OUT]	MovieListのIDが振られた個所にstrStartHtmlの内容を書き込む
'***********************************************************
	Sub ReadMovieFolder(obj,flg)
		Set objFileSys = CreateObject("Scripting.FileSystemObject")
		'初期画面からの遷移時のみチェック
		if flg = 1then
			if  objFileSys.FolderExists(baseFolder & obj) =false then
				msgbox("該当のフォルダが存在しません")
				exit Sub
			End if
		End If
	 	Set objFolder = objFileSys.GetFolder(baseFolder & obj)

		'取得して配列に格納する
		Dim fileList()
		if flg = 0 then
			cnt =0
			ReDim fileList(objFolder.Files.Count)
			For Each objItem In objFolder.Files
				fileList(cnt) = objItem.Name
				cnt=cnt +1
			Next
		Else
			cnt =0
			ReDim fileList(objFolder.SubFolders.Count)
			For Each objItem In objFolder.SubFolders
				fileList(cnt) = objItem.Name
				cnt=cnt +1
			Next
		End If

		'ソートする
		if flg =0 then	
			call QuickSort(fileList,0,objFolder.Files.Count)
			targetHtml =makehtmlFile(fileList,obj)
		Else
			call QuickSort(fileList,0,objFolder.SubFolders.Count)
			targetHtml =makehtmlFolder(fileList,obj)
		End If

		call readHtml("MovieList",targetHtml)
	End Sub


'***********************************************************
'取得した配列でhtmlを生成する（フォルダ一覧）
'[IN]	folderList	該当フォルダ名の配列
'[IN]	obj		上位の固定フォルダ文字列
'[OUT]	フォルダ一覧の表示用html文字列
'***********************************************************
	Function makehtmlFolder(folderList,obj)
		Dim targetHtml
		Set objFileSys = CreateObject("Scripting.FileSystemObject")

		targetHtml = "<table width=900px>TOPメニュー > " & obj & "</table><table><tr>"
		tr_count=1
		For Each folderName In folderList
			'イメージファイルがある場合それを設定
			imgData = "img/NoImage.jpg"
			 If objFileSys.FileExists("img/" & folderName & ".jpg") Then
				imgData = "img/" & folderName & ".jpg"
			End If

			if folderName ="" then
				
			else
				targetHtml = targetHtml & "<td>" & _
				"<table border=1 width=180 height=130><tr><td valign=top>" & _
				"<a href='#' onClick='ReadMovieFolder " & chr(34) & obj & "\" & folderName  & chr(34) & "," & chr(34) & "0" & chr(34) & "'>" & folderName  & "</a>" & _
				"<table class=thumbcell><tr><td valign=top  bgcolor='black'><div class=thumbleft><div class=thumbimage_small >" & _
				"<table class=thumbimagetb cellspacing=0 cellpadding=0><tr><td>" & _
				"<img src='" & imgData & "' width='80' align='LEFT'></td></tr></table>" & _
				"</div></div></td></tr></table>" & _
				"</td></tr></table>" & _
				"</td>"
				if tr_count =5 then
					targetHtml = targetHtml & "</tr><tr>"
					tr_count=0
				end If
				tr_count=tr_count +1
			end If
		Next
		'幅を一定にする為の処理
		if tr_count < 5 and tr_count > 1 then
			For i=0 to 5 - tr_count
				targetHtml = targetHtml & "<td>" & _
				"<table border=0 width=180 height=130><tr><td></td></tr></table>" & _
				"</td>"
			Next
		targetHtml = targetHtml & "</tr>"
		End if

		targetHtml = targetHtml &  "</table>" & _
		"<p><table  width=900px border =0><tr>" & _
		"<td align='right'><input class ='nomal' type ='button' value='TOPに戻る' onClick='startDisplay()'>　" & _
		"<input class ='nomal' type ='button' value='終了' onClick='window.close()'></td></tr></table>"
		makehtmlFolder= targetHtml
	End Function

'***********************************************************
'取得した配列でhtmlを生成する（ファイル一覧）
'[IN]	fileList		該当ファイル名の配列
'[IN]	obj		上位のフォルダ文字列
'[OUT]	ファイル一覧の表示用html文字列
'***********************************************************
	Function makehtmlFile(fileList,obj)

		Dim targetHtml
		Set objFileSys = CreateObject("Scripting.FileSystemObject")

		Dim returnKey
		returnKey = left(obj,1)
		word2 = replace(obj,returnKey,"")
		word2 = replace(word2,"\","")

		targetHtml = "<table width=900px>TOPメニュー > " & returnKey & " > " & word2 & "</table><table><tr>"
		tr_count=1
		For Each fileName In fileList
			'イメージファイルがある場合それを設定
			imgData = "img/NoImage.jpg"
			 If objFileSys.FileExists("img/" & fileName & ".jpg") Then
				imgData = "img/" & fileName & ".jpg"
			End If

			if fileName ="" then
				
			else
				targetHtml = targetHtml & "<td>" & _
				"<table border=1 width=180 height=130><tr><td valign=top>" & _
				"<a href='#' onClick='ExeMovie(" & chr(34) &  baseFolder & obj & "\" & fileName  & chr(34) & ")'>" & fileName  & "</a>" & _
				"<table class=thumbcell><tr><td valign=top  bgcolor='black'><div class=thumbleft><div class=thumbimage_small >" & _
				"<table class=thumbimagetb cellspacing=0 cellpadding=0><tr><td>" & _
				"<img src='" & imgData & "' width='80' align='LEFT'></td></tr></table>" & _
				"</div></div></td></tr></table>" & _
				"</td></tr></table>" & _
				"</td>"
				if tr_count =5 then
					targetHtml = targetHtml & "</tr><tr>"
					tr_count=0
				end If
				tr_count=tr_count +1
			end If
		Next
		'幅を一定にする為の処理
		if tr_count < 5  and tr_count > 1 then
			For i=0 to 5 - tr_count
				targetHtml = targetHtml & "<td>" & _
				"<table border=0 width=180 height=130><tr><td></td></tr></table>" & _
				"</td>"
			Next
			targetHtml = targetHtml & "</tr>"

		End if

		targetHtml = targetHtml &  "</table>" & _
		"<p><table  width=900px border =0><tr>" & _
		"<td align='right'><input class ='nomal' type ='button' value='一覧に戻る' onClick='ReadMovieFolder " & chr(34) &  returnKey & chr(34) & "," & chr(34) & "1" & chr(34) & "'>　" & _
		"<input class ='nomal' type ='button' value='終了' onClick='window.close()'></td></tr></table>"

		makehtmlFile= targetHtml
	End Function

'***********************************************************
'検索条件に合致するファイル一覧用htmlの作成
'[IN]	fileList		該当ファイル名の配列
'[OUT]	ファイル一覧の表示用html文字列
'***********************************************************
	Function makehtmlFileSearch(fileList)

		Dim targetHtml
		Set objFileSys = CreateObject("Scripting.FileSystemObject")

		targetHtml = "<table width=900px>TOPメニュー > 検索結果 </table><table><tr>"
		tr_count=1
		For Each fileName In fileList
			'パス抜きのファイル名を取得する
			FileNameOnly = getFileName(fileName)
			'イメージファイルがある場合それを設定(後で修正)
			imgData = "img/NoImage.jpg"
			 If objFileSys.FileExists("img/" & FileNameOnly & ".jpg") Then
				imgData = "img/" & FileNameOnly & ".jpg"
			End If

			if fileName ="" then
				
			else
				targetHtml = targetHtml & "<td>" & _
				"<table border=1 width=180 height=130><tr><td valign=top>" & _
				"<a href='#' onClick='ExeMovie(" & chr(34) &  baseFolder & fileName  & chr(34) & ")'>" & FileNameOnly  & "</a>" & _
				"<table class=thumbcell><tr><td valign=top  bgcolor='black'><div class=thumbleft><div class=thumbimage_small >" & _
				"<table class=thumbimagetb cellspacing=0 cellpadding=0><tr><td>" & _
				"<img src='" & imgData & "' width='80' align='LEFT'></td></tr></table>" & _
				"</div></div></td></tr></table>" & _
				"</td></tr></table>" & _
				"</td>"
				if tr_count =5 then
					targetHtml = targetHtml & "</tr><tr>"
					tr_count=0
				end If
				tr_count=tr_count +1
			end If
		Next
		'幅を一定にする為の処理
		if tr_count < 5  and tr_count > 1 then
			For i=0 to 5 - tr_count
				targetHtml = targetHtml & "<td>" & _
				"<table border=0 width=180 height=130><tr><td></td></tr></table>" & _
				"</td>"
			Next
			targetHtml = targetHtml & "</tr>"

		End if

		targetHtml = targetHtml &  "</table>" & _
		"<p><table  width=900px border =0><tr>" & _
		"<td align='right'><input class ='nomal' type ='button' value='TOPに戻る' onClick='startDisplay()'>　" & _
		"<input class ='nomal' type ='button' value='終了' onClick='window.close()'></td></tr></table>"

		makehtmlFileSearch= targetHtml
	End Function
'***********************************************************
'検索条件に合致するファイル名の配列を生成する
'[IN]	なし
'[OUT]	ファイル一覧の表示用html文字列
'***********************************************************
	Sub ReadSearchFiles()
	
	//全ファイルリストを取得
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFileSys.GetFolder(baseFolder)
		Dim fileList
		Dim searchList
	'ベース直下のフォルダすべてを取得
		ReDim fileList(1000)
		cnt =0
		For Each objItem In objFolder.SubFolders
			Set objSubFolder = objFileSys.GetFolder(baseFolder & objItem.Name)
			'直下のフォルダー以下のサブフォルダーを取得
			For Each objItem2 In objSubFolder.SubFolders
				Set objSubFolder2 = objFileSys.GetFolder(baseFolder & objItem.Name & "\" & objItem2.Name)
			'さらに下のファイル一覧を取得
				For Each objItem3 In objSubFolder2.Files
					fileList(cnt) = objItem.Name & "\" & objItem2.Name & "\" & objItem3.Name
					cnt = cnt + 1
				Next
			Next
		Next
	Dim keyword
	'画面側のオブジェクトからキーワードを取得する
	keyword= document.getElementById("searchText").value
	'検索する
	searchList = Filter(fileList,keyword,true,1)

	if Ubound(searchList) =0 or Ubound(searchList) = -1 then
		msgbox "検索条件に合致するファイルがありません"
		Exit Sub
	End If
	'htmlの生成
	targetHtml =makehtmlFileSearch(searchList)

	'画面書き換え
	call readHtml("MovieList",targetHtml)
	End Sub

'***********************************************************
'対象のIDをobjの内容で書き換える
'[IN]	targetId		書き換える対象のID名称
'[IN]	obj		書き換える内容
'[OUT]	対象のIDが指定されたhtmlをobjの内容に書き換える
'***********************************************************
	Sub readHtml(targetId,obj)
		document.getElementById(targetId).innerHTML =obj
	End Sub

'***********************************************************
'配列をソートする
'[IN]	vntArray		配列
'	vntStart		ソート対象の開始番号
'	vntEnd		ソート対象の終了番号
'[OUT]	vntArray	の内容をソートする
'***********************************************************
Sub QuickSort(vntArray, vntStart, vntEnd)

 Dim vntBaseNumber                                      '中央の要素番号を格納する変数
 Dim vntBaseValue                                       '基準値を格納する変数
 Dim vntCounter                                         '格納位置カウンタ
 Dim vntBuffer                                          '値をスワップするための作業域
 Dim i                                                  'ループカウンタ

If vntStart >= vntEnd Then Exit Sub                 '終了番号が開始番号以下の場合、プロシージャを抜ける

vntBaseNumber = (vntStart + vntEnd) \ 2             '中央の要素番号を求める
vntBaseValue = vntArray(vntBaseNumber)              '中央の値を基準値とする
vntArray(vntBaseNumber) = vntArray(vntStart)        '中央の要素に開始番号の値を格納
vntCounter = vntStart                               '格納位置カウンタを開始番号と同じにする

For i = (vntStart + 1) To vntEnd Step 1             '開始番号の次の要素から終了番号までループ
	If vntArray(i) < vntBaseValue Then              '値が基準値より小さい場合
		vntCounter = vntCounter + 1                 '格納位置カウンタをインクリメント
		vntBuffer = vntArray(vntCounter)            'vntArray(i) と vntArray(vntCounter) の値をスワップ
		vntArray(vntCounter) = vntArray(i)
		vntArray(i) = vntBuffer
	End If
Next
vntArray(vntStart) = vntArray(vntCounter)           'vntArray(vntCounter) を開始番号の値にする
vntArray(vntCounter) = vntBaseValue                 '基準値を vntArray(vntCounter) に格納
Call QuickSort(vntArray, vntStart, vntCounter - 1)  '分割された配列をクイックソート(再帰)
Call QuickSort(vntArray, vntCounter + 1, vntEnd)    '分割された配列をクイックソート(再帰)

End Sub

'***********************************************************
'動画ファイルを実行する
'[IN]	obj　該当ファイル名
'[OUT]関連付けられたPGでファイルを実行する
'***********************************************************
Sub ExeMovie(obj)
	Set objShell = CreateObject("Wscript.Shell") 
	objShell.Run  chr(34) & obj & chr(34)
	
End Sub

'***********************************************************
'ファイル名を取得する
'[IN]	obj　該当ファイル名(パス付)
'[OUT]　パスを外したファイル名
'***********************************************************
Function getFileName(obj)
	NameSize = Len(obj)

	For y = 1 To NameSize 
		strtmp = Right(obj, y)
		strtmp = Left(strtmp, 1)
		If strtmp = "\" Then
			y = y - 1
		    	GetFileExt = Right(obj, y)
			Exit For
		End If
	Next
	getFileName = GetFileExt
End Function