'***********************************************************
' * All Rights Reserved, Copyright(C) korai 2015
' * �v���O�������F  ����Ǘ��V�X�e��
' * �v���O�����h�c�Fconfig.vbs
' * �쐬���t�F	   2015/04/30
' * �ŏI�X�V���F	   2015/04/30
' * �v���O���������F�萔���W���[��
' * �ύX����
' * �ύX����		�ύX�Җ�			����
' * 2015/04/30	�g��				�V�K�쐬
'************************************************************

'����t�@�C���̃��[�g

baseFolder =targetPath & "E:\�A�j��\"

'�A�v���̃T�C�Y
appWidth =980
appHeight=800

'�N���ꏊ�̎w��i�����j
posx = (screen.width -appWidth) /2
posy = (screen.height -appHeight) /2
window.moveTo posx, posy

'�E�B���h�E�T�C�Y�̎w��
window.resizeTo appWidth,appHeight

