'***********************************************************
' * All Rights Reserved, Copyright(C) korai 2015
' * �v���O�������F  ����Ǘ��V�X�e��
' * �v���O�����h�c�Fcommon.vbs
' * �쐬���t�F	   2015/04/30
' * �ŏI�X�V���F	   2015/04/30
' * �v���O���������F���ʃ��W���[��
' * �ύX����
' * �ύX����		�ύX�Җ�			����
' * 2015/04/30	�g��				�V�K�쐬
'************************************************************

'***********************************************************
'������ʂ�html���o�͂���
'[IN]	�Ȃ�
'[OUT]	MovieList��ID���U��ꂽ����strStartHtml�̓��e����������
'***********************************************************
	Sub startDisplay()
		Dim strStartHtml
		strStartHtml ="<table width='85%'><tr><td class='top'>TOP���j���[</td>" & _
		"<td align='right'><input class='search' type='text' id='searchText'>" & _
		"<input class ='searchbtn' type='button' value='����' onClick='ReadSearchFiles()'>" & _
		"</td></tr></table><p>"

		strStartHtml = strStartHtml & "<table><tr>" & _
		"<td><input  type ='button' value='��' onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  class ='tile' type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  class ='tile' type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"</tr>" & _
		"<tr>" & _
		"<td><input  class ='tile' type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  class ='tile' type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"<td><input  class ='tile' type ='button' value='��'  onClick='ReadMovieFolder " & chr(34) & "��" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"</tr>" & _
		"<tr>" & _
		"<td><input  type ='button' value='���̑�'  onClick='ReadMovieFolder " & chr(34) & "���̑�" &  chr(34) &"," & chr(34) & "1" & chr(34) & "'></td>" & _
		"</tr></table>" & _
		"<p><table  width=750px ><tr><td align='right'><input class ='nomal' type ='button' value='�I��' onClick='window.close()'></td></tr></table>"

		call readHtml("MovieList",strStartHtml)
	End Sub
'***********************************************************
'�t�H���_�ǂݍ���ňȉ��ɑ��݂���t�@�C���A�������̓t�H���_����html���o�͂���
'[IN]	obj�@�Y���t�H���_��
'[IN]	flg�@0�F�t�@�C���@1�F�t�H���_
'[OUT]	MovieList��ID���U��ꂽ����strStartHtml�̓��e����������
'***********************************************************
	Sub ReadMovieFolder(obj,flg)
		Set objFileSys = CreateObject("Scripting.FileSystemObject")
		'������ʂ���̑J�ڎ��̂݃`�F�b�N
		if flg = 1then
			if  objFileSys.FolderExists(baseFolder & obj) =false then
				msgbox("�Y���̃t�H���_�����݂��܂���")
				exit Sub
			End if
		End If
	 	Set objFolder = objFileSys.GetFolder(baseFolder & obj)

		'�擾���Ĕz��Ɋi�[����
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

		'�\�[�g����
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
'�擾�����z���html�𐶐�����i�t�H���_�ꗗ�j
'[IN]	folderList	�Y���t�H���_���̔z��
'[IN]	obj		��ʂ̌Œ�t�H���_������
'[OUT]	�t�H���_�ꗗ�̕\���phtml������
'***********************************************************
	Function makehtmlFolder(folderList,obj)
		Dim targetHtml
		Set objFileSys = CreateObject("Scripting.FileSystemObject")

		targetHtml = "<table width=900px>TOP���j���[ > " & obj & "</table><table><tr>"
		tr_count=1
		For Each folderName In folderList
			'�C���[�W�t�@�C��������ꍇ�����ݒ�
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
		'�������ɂ���ׂ̏���
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
		"<td align='right'><input class ='nomal' type ='button' value='TOP�ɖ߂�' onClick='startDisplay()'>�@" & _
		"<input class ='nomal' type ='button' value='�I��' onClick='window.close()'></td></tr></table>"
		makehtmlFolder= targetHtml
	End Function

'***********************************************************
'�擾�����z���html�𐶐�����i�t�@�C���ꗗ�j
'[IN]	fileList		�Y���t�@�C�����̔z��
'[IN]	obj		��ʂ̃t�H���_������
'[OUT]	�t�@�C���ꗗ�̕\���phtml������
'***********************************************************
	Function makehtmlFile(fileList,obj)

		Dim targetHtml
		Set objFileSys = CreateObject("Scripting.FileSystemObject")

		Dim returnKey
		returnKey = left(obj,1)
		word2 = replace(obj,returnKey,"")
		word2 = replace(word2,"\","")

		targetHtml = "<table width=900px>TOP���j���[ > " & returnKey & " > " & word2 & "</table><table><tr>"
		tr_count=1
		For Each fileName In fileList
			'�C���[�W�t�@�C��������ꍇ�����ݒ�
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
		'�������ɂ���ׂ̏���
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
		"<td align='right'><input class ='nomal' type ='button' value='�ꗗ�ɖ߂�' onClick='ReadMovieFolder " & chr(34) &  returnKey & chr(34) & "," & chr(34) & "1" & chr(34) & "'>�@" & _
		"<input class ='nomal' type ='button' value='�I��' onClick='window.close()'></td></tr></table>"

		makehtmlFile= targetHtml
	End Function

'***********************************************************
'���������ɍ��v����t�@�C���ꗗ�phtml�̍쐬
'[IN]	fileList		�Y���t�@�C�����̔z��
'[OUT]	�t�@�C���ꗗ�̕\���phtml������
'***********************************************************
	Function makehtmlFileSearch(fileList)

		Dim targetHtml
		Set objFileSys = CreateObject("Scripting.FileSystemObject")

		targetHtml = "<table width=900px>TOP���j���[ > �������� </table><table><tr>"
		tr_count=1
		For Each fileName In fileList
			'�p�X�����̃t�@�C�������擾����
			FileNameOnly = getFileName(fileName)
			'�C���[�W�t�@�C��������ꍇ�����ݒ�(��ŏC��)
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
		'�������ɂ���ׂ̏���
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
		"<td align='right'><input class ='nomal' type ='button' value='TOP�ɖ߂�' onClick='startDisplay()'>�@" & _
		"<input class ='nomal' type ='button' value='�I��' onClick='window.close()'></td></tr></table>"

		makehtmlFileSearch= targetHtml
	End Function
'***********************************************************
'���������ɍ��v����t�@�C�����̔z��𐶐�����
'[IN]	�Ȃ�
'[OUT]	�t�@�C���ꗗ�̕\���phtml������
'***********************************************************
	Sub ReadSearchFiles()
	
	//�S�t�@�C�����X�g���擾
	Set objFileSys = CreateObject("Scripting.FileSystemObject")
	Set objFolder = objFileSys.GetFolder(baseFolder)
		Dim fileList
		Dim searchList
	'�x�[�X�����̃t�H���_���ׂĂ��擾
		ReDim fileList(1000)
		cnt =0
		For Each objItem In objFolder.SubFolders
			Set objSubFolder = objFileSys.GetFolder(baseFolder & objItem.Name)
			'�����̃t�H���_�[�ȉ��̃T�u�t�H���_�[���擾
			For Each objItem2 In objSubFolder.SubFolders
				Set objSubFolder2 = objFileSys.GetFolder(baseFolder & objItem.Name & "\" & objItem2.Name)
			'����ɉ��̃t�@�C���ꗗ���擾
				For Each objItem3 In objSubFolder2.Files
					fileList(cnt) = objItem.Name & "\" & objItem2.Name & "\" & objItem3.Name
					cnt = cnt + 1
				Next
			Next
		Next
	Dim keyword
	'��ʑ��̃I�u�W�F�N�g����L�[���[�h���擾����
	keyword= document.getElementById("searchText").value
	'��������
	searchList = Filter(fileList,keyword,true,1)

	if Ubound(searchList) =0 or Ubound(searchList) = -1 then
		msgbox "���������ɍ��v����t�@�C��������܂���"
		Exit Sub
	End If
	'html�̐���
	targetHtml =makehtmlFileSearch(searchList)

	'��ʏ�������
	call readHtml("MovieList",targetHtml)
	End Sub

'***********************************************************
'�Ώۂ�ID��obj�̓��e�ŏ���������
'[IN]	targetId		����������Ώۂ�ID����
'[IN]	obj		������������e
'[OUT]	�Ώۂ�ID���w�肳�ꂽhtml��obj�̓��e�ɏ���������
'***********************************************************
	Sub readHtml(targetId,obj)
		document.getElementById(targetId).innerHTML =obj
	End Sub

'***********************************************************
'�z����\�[�g����
'[IN]	vntArray		�z��
'	vntStart		�\�[�g�Ώۂ̊J�n�ԍ�
'	vntEnd		�\�[�g�Ώۂ̏I���ԍ�
'[OUT]	vntArray	�̓��e���\�[�g����
'***********************************************************
Sub QuickSort(vntArray, vntStart, vntEnd)

 Dim vntBaseNumber                                      '�����̗v�f�ԍ����i�[����ϐ�
 Dim vntBaseValue                                       '��l���i�[����ϐ�
 Dim vntCounter                                         '�i�[�ʒu�J�E���^
 Dim vntBuffer                                          '�l���X���b�v���邽�߂̍�ƈ�
 Dim i                                                  '���[�v�J�E���^

If vntStart >= vntEnd Then Exit Sub                 '�I���ԍ����J�n�ԍ��ȉ��̏ꍇ�A�v���V�[�W���𔲂���

vntBaseNumber = (vntStart + vntEnd) \ 2             '�����̗v�f�ԍ������߂�
vntBaseValue = vntArray(vntBaseNumber)              '�����̒l����l�Ƃ���
vntArray(vntBaseNumber) = vntArray(vntStart)        '�����̗v�f�ɊJ�n�ԍ��̒l���i�[
vntCounter = vntStart                               '�i�[�ʒu�J�E���^���J�n�ԍ��Ɠ����ɂ���

For i = (vntStart + 1) To vntEnd Step 1             '�J�n�ԍ��̎��̗v�f����I���ԍ��܂Ń��[�v
	If vntArray(i) < vntBaseValue Then              '�l����l��菬�����ꍇ
		vntCounter = vntCounter + 1                 '�i�[�ʒu�J�E���^���C���N�������g
		vntBuffer = vntArray(vntCounter)            'vntArray(i) �� vntArray(vntCounter) �̒l���X���b�v
		vntArray(vntCounter) = vntArray(i)
		vntArray(i) = vntBuffer
	End If
Next
vntArray(vntStart) = vntArray(vntCounter)           'vntArray(vntCounter) ���J�n�ԍ��̒l�ɂ���
vntArray(vntCounter) = vntBaseValue                 '��l�� vntArray(vntCounter) �Ɋi�[
Call QuickSort(vntArray, vntStart, vntCounter - 1)  '�������ꂽ�z����N�C�b�N�\�[�g(�ċA)
Call QuickSort(vntArray, vntCounter + 1, vntEnd)    '�������ꂽ�z����N�C�b�N�\�[�g(�ċA)

End Sub

'***********************************************************
'����t�@�C�������s����
'[IN]	obj�@�Y���t�@�C����
'[OUT]�֘A�t����ꂽPG�Ńt�@�C�������s����
'***********************************************************
Sub ExeMovie(obj)
	Set objShell = CreateObject("Wscript.Shell") 
	objShell.Run  chr(34) & obj & chr(34)
	
End Sub

'***********************************************************
'�t�@�C�������擾����
'[IN]	obj�@�Y���t�@�C����(�p�X�t)
'[OUT]�@�p�X���O�����t�@�C����
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