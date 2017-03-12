'***********************************************************
' * All Rights Reserved, Copyright(C) korai 2015
' * プログラム名：  動画管理システム
' * プログラムＩＤ：config.vbs
' * 作成日付：	   2015/04/30
' * 最終更新日：	   2015/04/30
' * プログラム説明：定数モジュール
' * 変更履歴
' * 変更日時		変更者名			説明
' * 2015/04/30	紅雷				新規作成
'************************************************************

'動画ファイルのルート

baseFolder =targetPath & "E:\アニメ\"

'アプリのサイズ
appWidth =980
appHeight=800

'起動場所の指定（中央）
posx = (screen.width -appWidth) /2
posy = (screen.height -appHeight) /2
window.moveTo posx, posy

'ウィンドウサイズの指定
window.resizeTo appWidth,appHeight

