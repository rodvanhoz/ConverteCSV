Option Explicit

' constantes
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

' variaveis de arquivos
Dim fs, arqent, arqsai, arqlay, arqlog

' variaveis hash
Dim hasharqent, hasharqsai, lay, hasharqlog

' variaveis diversas da aplicação
Dim linha, linhasai, cont, campos, x, y, narqent, narqsai, narqlog, nlay, totalcampo, linhatkn, sep

' checando total de parametros
If WScript.Arguments.Count <> 4 Then
	msgErro "USO: ConverteCSV [Layout] [ArqEntrada] [ArqSaida] [Delimitador]"
End if

' parametros
narqent = WScript.Arguments.Unnamed(1)
narqsai = WScript.Arguments.Unnamed(2)
narqlog = "ConverteCSV.LOG"
nlay    = WScript.Arguments.Unnamed(0)

' definindo separador
If WScript.Arguments.Unnamed(3) = 1 Then
	sep = ";"
ElseIf WScript.Arguments.Unnamed(3) = 2 Then
	sep = "|"
Else
	msgErro "ERRO: Delimitador invalido. 1 = ; // 2 = |"
End If

' setando variaveis
Set fs = CreateObject( "scripting.FileSystemObject" )
Set arqent = fs.OpenTextFile( narqent, ForReading, TristateFalse )
Set arqsai = fs.CreateTextFile( narqsai )

' hash
Set hasharqent = CreateObject( "Scripting.Dictionary" )
Set hasharqsai = CreateObject( "Scripting.Dictionary" )
Set hasharqlog = CreateObject( "Scripting.Dictionary" )
Set lay = CreateObject( "Scripting.Dictionary" )

' definindo valores das variaveis
cont = 0

' INICIO DA IMPLEMENTAÇÃO
	carregaLay( nlay )
	totalcampo = lay.Count
	
	Do Until arqent.AtEndOfStream
		linhasai = Empty
		linha = arqent.ReadLine
		linhatkn = Split( linha, sep )
		cont = cont + 1
		
		campos = lay.Keys
		
		If UBound(campos) > UBound(linhatkn) Then
			msgErro "ERRO: Total de campos no layout é superior a que o total de campos da linha " & CStr( cont )
		End if
		
		For x = 0 To UBound(linhatkn)
			If UBound(campos) < x Then
				msgErro "ERRO: Total de campos no layout e menor que total de campos da linha " & CStr( cont ) & " - Campo: " & linhatkn(x)
			End if
		
			linhasai = linhasai & acerta( linhatkn(x), getTamLay( campos(x) ), "E", campos(x) )
		Next
		
		arqsai.WriteLine linhasai
		WScript.StdOut.Write Chr(13) & "Registros: " & CStr( cont )
	Loop
	
	arqent.Close
	arqsai.Close
	
	WScript.StdOut.WriteBlankLines 2
	WScript.Quit(0)
	
' FINAL DA IMPLEMENTAÇÃO

' FUNÇÕES
Sub carregaLay( caminho )
	Dim linha, tkn, tamlinha
	
	Set arqlay = fs.OpenTextFile( caminho, ForReading, TristateFalse )
	tamlinha = 1
	
	Do Until arqlay.AtEndOfStream
		linha = arqlay.ReadLine
		tkn = Split( linha, ";" )
		
		If lay.Exists( LCase(tkn(0)) ) Then
			msgErro "ERRO: Campo " * tkn(0) & " ja existe no layout"
		End If
		
		lay.Add LCase(tkn(0)), tkn(1) & ";" & CStr( tamlinha )
		
		tamlinha = tamlinha & CInt( tkn(1) )
	Loop
	
	arqlay.Close
End Sub

Function getTamLay( nomecampo )
	Dim tkn
	
	If Not lay.Exists( LCase(nomecampo) ) Then
		msgErro "ERRO: Campo " & nomecampo & " nao existe no layout"
	End If
	
	tkn = Split( lay.Item( LCase(nomecampo) ), ";" )
	getTamLay = tkn(0)
End function

Function getPosLay( nomecampo )
	Dim tkn
	
	If Not lay.Exists( LCase(nomecampo) ) Then
		msgErro "ERRO: Campo " & nomecampo & " nao existe no layout"
	End If
	
	tkn = Split( lay.Item( LCase(nomecampo) ), ";" )
	getPosLay = tkn(1)
End Function

Function getLinhaLay( linha, nomecampo )
	Dim pos, tam
	
	If Not lay.Exists( LCase(nomecampo) ) Then
		msgErro "ERRO: Campo " & nomecampo & " nao existe no layout"
	End If
	
	pos = getPosLay( nomecampo )
	tam = getTamLay( nomecampo )
	
	If pos > Len( linha ) Then
		msgErro "ERRO: posicao " & CStr( pos ) & " excedeu o tamanho da linha. Tamanho: " & CStr( tam )
	End If
	
	getLinhaLay = Mid( linha, pos, tam )
End function

Function acerta( cont, tam, algn, nomecampo )
	Dim tmpE, tmpD, x
	
	If Len( cont ) > CINT(tam) Then
		msgErro "ERRO: Conteudo e maior que o campo " & nomecampo
	Else
		For x = 1 To (tam - Len( cont ))
			If x < ( (tam - Len( cont )) / 2 ) Then
				tmpE = tmpE & " "
			Else
				tmpD = tmpD & " "
			End If
		Next
		
		If algn = "E" Then
			acerta = cont & tmpE & tmpD
		ElseIf algn = "D" Then
			acerta = tmpE & tmpD & cont
		ElseIf algn = "D" Then
			acerta = tmpE & cont & tmpD
		Else
			msgErro "ERRO: Alinhamento nao previsto: " & algn
		End if
	End if
End function

Sub msg( m )
	WScript.StdOut.WriteBlankLines 1
	WScript.StdOut.WriteLine m
End sub

Sub msgErro( m )
	msg m
	WScript.StdOut.WriteBlankLines 1
	
	Set arqlog = fs.CreateTextFile( narqlog )
	arqlog.WriteLine m
	arqlog.Close
	
	WScript.Quit( 1 )
End Sub