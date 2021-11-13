REM  *****  BASIC  *****

Sub CreateTemplateSheet(pOption as String, Optional pQty as Integer, Optional pLang as String)
dim vLang as String
Dim firstName(1 to 100) As String
Dim lastName(1 to 50) As String
dim course(1 to 20) as String
dim uf (1 to 27) as String
dim vNewName as String
dim vQty as Integer
Dim vNumberName as Integer
Dim vNumberLastName as Integer


firstName = Array("Adriana","Adriano","Alan","Albano","Aldo","Alexandre","Alice","Alice","Aline","Ana","Anderson","André","Andressa","Angélica","Antônia","Antônio","Arthur","Auri","Barbara","Beatriz","Benício","Bernardo","Caio","Camila","Carina","Carlos","Carolina","Cecília","Cecília","Celso","Cícero","Cláudia","Cláudio","Daniel","Davi","Diana","Eduardo","Emanuel","Enzo Gabriel","Fabiana","Fábio","Felipe","Fernanda","Fernando","Francisco","Gabriel","Gael","Gilberto","Gustavo","Heitor","Helena","Heloísa","Hugo","Isabella","Jaqueline","Jeferson","João","João","João Miguel","Júlia","Júlio","Juracy","Larissa","Laura","Leila","Leonardo","Letícia","Lívia","Lorena","Lorena","Lorenzo","Luan","Luana","Luana","Lucas","Luiz","Luiza","Manuel","Manuela","Marcelo","Marcos","Maria Cecília","Maria Clara","Maria Eduarda","Maria Júlia","Mariana","Marina","Maurício","Michelle","Miguel","Patrícia","Pedro","Samuel","Sophia","Theo","Thiago","Valentina","Vanessa","Viviane","Zenaide")
lastName = Array("Albuquerque","Almeida","Alves","Andrade","Barbosa","Barros","Barroso","Batista","Borges","Cabanas","Campos","Cardoso","Castro","Cavalcanti","Cista","Conceição","Costa","Cremer","Dias","Duarte","Esteves","Fernandes","Ferreira","Freitas","Garcia","Gomes","Gonçalves","Lima","Lopes","Machado","Malheiros","Marques","Martins","Mendes","Miranda","Moraes","Moreira","Nascimento","Nunes","Oliveira","Pereira","Ribeiro","Rodrigues","Santana","Santos","Silva","Soares","Souza","Teixeira","Vieira")
course = Array("Arquitetura","Astronomia","Biologia","Ciências Sociais","Direito", "Educação Artísitca","Educação Física", "Enfermagem","Engenharia","Física","História","Informática", "Letras","Matemática", "Medicina", "Pedagogia", "Psicologia", "Química","Relações Internacionais","Serviço Social")
uf = Array("Acre","Alagoas","Amapá","Amazonas","Bahia","Ceará","Espírito Santo","Goiás","Maranhão","Mato Grosso","Mato Grosso do Sul","Minas Gerais","Pará","Paraíba","Paraná","Pernambuco","Piauí","Rio de Janeiro","Rio Grande do Norte","Rio Grande do Sul","Rondônia","Roraima","Santa Catarina","São Paulo","Sergipe","Tocantins","Distrito Federal")

	if IsMissing(pQty) Then
			vQty = 100	
		Else
			vQty = pQty
	end If
	
	if IsMissing(pLang) Then
			vLang = "pt"	
		Else
			vLang = pLang
	end If
	
	if pOption = "uf" then
	
		CreateSheet("UF")
	
			For i= 2 To 28 Step 1
			
				Cell("UF", REF(i, 1)).String = uf ( i - 2 )
				
			next
			
			Cell ("UF", "A1").String = "Unidades Federativas (BR)"
			
			SortAsc("UF", "A2:A30",1)
			
			sheet("UF").getColumns().getByIndex(0).OptimalWidth = True
	
	end if


	if pOption = "names" or pOption = "students" then
	
		CreateSheet("Nomes")
		
		Cell("Nomes", "A1").String = "Nome"

		For i= 2 To vQty Step 1
		
			Do 
	   		
	   			vNumberName = ( Int ( Rnd * 99) + 1 )
				vNumberLastName = ( Int ( Rnd * 49) + 1 )
				vNewName = firstName( vNumberName ) & " " & lastName( vNumberLastName)
		
			Loop until not FindTextInColumn(vNewName, "Nomes", 1, i+1)
		
			Cell("Nomes", REF(i, 1)).String = vNewName
			
		next
		
		SortAsc("Nomes", "A2:A10000", 1)		
		
		sheet("Nomes").getColumns().getByIndex(0).OptimalWidth = True
		
	end if
		
	if pOption = "students" then
	
		Cell("Nomes", "B1").String = "Unidade"
		Cell("Nomes", "C1").String = "Curso"
		
		For i= 2 To vQty Step 1
	
   			vNumberCourse = ( Int ( Rnd * 19) + 1 )
   			
   			vNumberUF =  ( Int ( Rnd * 26) + 1 )
		
			Cell("Nomes", REF(i, 2)).String = uf ( vNumberUF )
			Cell("Nomes", REF(i, 3)).String = Course ( vNumberCourse )
			
		next
			
		sheet("Nomes").getColumns().getByIndex(1).OptimalWidth = True
		sheet("Nomes").getColumns().getByIndex(2).OptimalWidth = True
	
	end if

End Sub
