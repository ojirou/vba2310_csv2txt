Attribute VB_Name = "Module1"
'#############################################################################
' CSV�t�@�C����TXT�t�@�C���ɕϊ�
'�@convert_csv2txt
'#############################################################################
Sub ConvertCSVtoTXT()
    Dim OutputCsv As String, OutputTxt As String
    Dim InputFileNumber As Integer, OutputFileNumber As Integer
    Dim LineText As String
    ' CSV�t�@�C���̃p�X��ݒ�
    OutputCsv = Environ("USERPROFILE") & "\Desktop\sample.csv"
    ' TXT�t�@�C���̃p�X��ݒ�
    OutputTxt = Environ("USERPROFILE") & "\Desktop\sample.txt"
    ' CSV�t�@�C�����e�L�X�g�t�@�C���ɕϊ�
    InputFileNumber = FreeFile
    Open OutputCsv For Input As #InputFileNumber
    OutputFileNumber = FreeFile
    Open OutputTxt For Output As #OutputFileNumber
    Do While Not EOF(InputFileNumber)
        Line Input #InputFileNumber, LineText
        ' �_�u���N�H�[�e�[�V����3�A�����_�u���N�H�[�e�[�V����1�ɒu��
        LineText = Replace(LineText, """""", "")
        ' �s���̋󔒂��폜
        LineText = Trim(LineText)
        ' �e�L�X�g�t�@�C���ɏ�������
        Print #OutputFileNumber, LineText
    Loop
    Close #InputFileNumber
    Close #OutputFileNumber
    MsgBox "CSV�t�@�C�����e�L�X�g�t�@�C���ɕϊ����܂����B", vbInformation
End Sub
