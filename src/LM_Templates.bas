REM  *****  BASIC  *****

'CreateTemplateSheet: Quickly generates template sheets with Brazilian states (UFs), random names, and (optionally) student data (State + Course).
'Useful for mocking data, testing formulas, charts, and workflows.
'
'pOption: Type of template to create (text) → "uf", "names" or "students"
'         "uf"       : Creates a sheet with the U.S. states (or BR states in pt)
'         "names"    : Creates a sheet with random unique first+last names
'         "students" : Same as "names", but also adds random "Unit" (State) and "Course"
'pQty: Number of rows to generate. Default = 100
'pLang: Language for the generated template sheets
'         "pt" : Uses Brazilian names, surnames, courses, and UFs
'         "en" : Uses U.S. names, surnames, courses, and States
Sub CreateTemplateSheet(pOption as String, Optional pLang as String, Optional pQty as Integer)
dim vLang as String
Dim firstName(1 to 100) As String
Dim lastName(1 to 50) As String
dim course(1 to 20) as String
dim uf (1 to 27) as String
dim vNewName as String
dim vQty as Integer
Dim vNumberName as Integer
Dim vNumberLastName as Integer
Dim vNameTitle as String
Dim vCouseTitle as String
Dim vCityTitle as String

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


if vLang = "pt" then

	firstName = Array("Adriana","Adriano","Alan","Albano","Aldo","Alexandre","Alice","Alice","Aline","Ana","Anderson","André","Andressa","Angélica","Antônia","Antônio","Arthur","Auri","Barbara","Beatriz","Benício","Bernardo","Caio","Camila","Carina","Carlos","Carolina","Cecília","Cecília","Celso","Cícero","Cláudia","Cláudio","Daniel","Davi","Diana","Eduardo","Emanuel","Enzo Gabriel","Fabiana","Fábio","Felipe","Fernanda","Fernando","Francisco","Gabriel","Gael","Gilberto","Gustavo","Heitor","Helena","Heloísa","Hugo","Isabella","Jaqueline","Jeferson","João","João","João Miguel","Júlia","Júlio","Juracy","Larissa","Laura","Leila","Leonardo","Letícia","Lívia","Lorena","Lorena","Lorenzo","Luan","Luana","Luana","Lucas","Luiz","Luiza","Manuel","Manuela","Marcelo","Marcos","Maria Cecília","Maria Clara","Maria Eduarda","Maria Júlia","Mariana","Marina","Maurício","Michelle","Miguel","Patrícia","Pedro","Samuel","Sophia","Theo","Thiago","Valentina","Vanessa","Viviane","Zenaide")
	lastName = Array("Albuquerque","Almeida","Alves","Andrade","Barbosa","Barros","Barroso","Batista","Borges","Cabanas","Campos","Cardoso","Castro","Cavalcanti","Cista","Conceição","Costa","Cremer","Dias","Duarte","Esteves","Fernandes","Ferreira","Freitas","Garcia","Gomes","Gonçalves","Lima","Lopes","Machado","Malheiros","Marques","Martins","Mendes","Miranda","Moraes","Moreira","Nascimento","Nunes","Oliveira","Pereira","Ribeiro","Rodrigues","Santana","Santos","Silva","Soares","Souza","Teixeira","Vieira")
	course = Array("Arquitetura","Astronomia","Biologia","Ciências Sociais","Direito", "Educação Artísitca","Educação Física", "Enfermagem","Engenharia","Física","História","Informática", "Letras","Matemática", "Medicina", "Pedagogia", "Psicologia", "Química","Relações Internacionais","Serviço Social")
	uf = Array("Acre","Alagoas","Amapá","Amazonas","Bahia","Ceará","Espírito Santo","Goiás","Maranhão","Mato Grosso","Mato Grosso do Sul","Minas Gerais","Pará","Paraíba","Paraná","Pernambuco","Piauí","Rio de Janeiro","Rio Grande do Norte","Rio Grande do Sul","Rondônia","Roraima","Santa Catarina","São Paulo","Sergipe","Tocantins","Distrito Federal")
	vNameTitle = "Nome"
	vCourseTitle = "Curso"
	vCityTitle = "Unidades Federativas (BR)"

else

	firstName = Array("James","John","Robert","Michael","William","David","Richard","Joseph","Thomas","Charles","Christopher","Daniel","Matthew","Anthony","Donald","Mark","Paul","Steven","Andrew","Joshua","Mary","Patricia","Jennifer","Linda","Elizabeth","Barbara","Susan","Jessica","Sarah","Karen","Lisa","Nancy","Betty","Margaret","Sandra","Ashley","Kimberly","Emily","Donna","Michelle","Carol","Amanda","Melissa","Deborah","Stephanie","Rebecca","Laura","Sharon","Cynthia","Kathleen","Amy","Angela","Anna","Ruth","Brenda","Pamela","Nicole","Katherine","Samantha","Christine","Janet","Maria","Heather","Diane","Julie","Joyce","Victoria","Kelly","Christina","Lauren","Joan","Evelyn","Olivia","Emma","Sophia","Isabella","Mia","Charlotte","Amelia","Harper","Evelyn","Abigail","Emily","Madison","Elizabeth","Sofia","Avery","Ella","Scarlett","Grace","Chloe","Benjamin","Lucas","Mason","Ethan","Alexander","Henry","Jacob","Jackson","Levi","Sebastian")
	lastName = Array("Smith","Johnson","Williams","Brown","Jones","Garcia","Miller","Davis","Rodriguez","Martinez","Hernandez","Lopez","Gonzalez","Wilson","Anderson","Thomas","Taylor","Moore","Jackson","Martin","Lee","Perez","Thompson","White","Harris","Sanchez","Clark","Ramirez","Lewis","Robinson","Walker","Young","Allen","King","Wright","Scott","Torres","Nguyen","Hill","Flores","Green","Adams","Nelson","Baker","Hall","Rivera","Campbell","Mitchell","Carter","Roberts")
	course = Array("Business Administration","Computer Science","Nursing","Psychology","Biology","Education","Accounting","Marketing","Mechanical Engineering","Electrical Engineering","Political Science","Economics","Sociology","History","English Literature","Chemistry","Mathematics","Criminal Justice","Communication Studies","Public Health","Civil Engineering","Philosophy","Art History","Environmental Science","Finance","Journalism","Anthropology","International Relations","Physics","Social Work")
	uf = Array("Alabama","Alaska","Arizona","Arkansas","California","Colorado","Connecticut","Delaware","Florida","Georgia","Hawaii","Idaho","Illinois","Indiana","Iowa","Kansas","Kentucky","Louisiana","Maine","Maryland","Massachusetts","Michigan","Minnesota","Mississippi","Missouri","Montana","Nebraska","Nevada","New Hampshire","New Jersey","New Mexico","New York","North Carolina","North Dakota","Ohio","Oklahoma","Oregon","Pennsylvania","Rhode Island","South Carolina","South Dakota","Tennessee","Texas","Utah","Vermont","Virginia","Washington","West Virginia","Wisconsin","Wyoming","District of Columbia")
	vNameTitle = "Name"
	vCourseTitle = "Course"
	vCityTitle = "City"

end if

	
	if pOption = "uf" and not SheetExists( "uf") then
	
		CreateSheet(pOption)
	
			For i= 2 To UBound(uf)+2 Step 1
			
				Cell(pOption, REF(i, 1)).String = uf ( i - 2 )
				
			next
			
			Cell (pOption, "A1").String = vCityTitle
			
			SortAsc(pOption, "A2:A30",1)
			
			sheet(pOption).getColumns().getByIndex(0).OptimalWidth = True
	
	end if


	if (pOption = "names" or pOption = "students") and not (SheetExists("names") or SheetExists("students"))  then
	
		CreateSheet(pOption)
		
		Cell(pOption, "A1").String = vNameTitle

		For i= 2 To vQty+1 Step 1
		
			Do 
	   		
	   			vNumberName = ( Int ( Rnd * 99) + 1 )
				vNumberLastName = ( Int ( Rnd * 49) + 1 )
				vNewName = firstName( vNumberName ) & " " & lastName( vNumberLastName)
		
			Loop until not FindTextInColumn(vNewName, pOption, 1, i+1)
		
			Cell(pOption, REF(i, 1)).String = vNewName
			
		next
		
		SortAsc(pOption, "A2:A10000", 1)		
		
		sheet(pOption).getColumns().getByIndex(0).OptimalWidth = True
		
	end if
		
	if pOption = "students" then
	
		Cell(pOption, "B1").String = vCityTitle
		Cell(pOption, "C1").String = vCourseTitle
		
		For i= 2 To vQty Step 1
	
   			vNumberCourse = ( Int ( Rnd * 19) + 1 )
   			
   			vNumberUF =  ( Int ( Rnd * 26) + 1 )
		
			Cell(pOption, REF(i, 2)).String = uf ( vNumberUF )
			Cell(pOption, REF(i, 3)).String = Course ( vNumberCourse )
			
		next
			
		sheet(pOption).getColumns().getByIndex(1).OptimalWidth = True
		sheet(pOption).getColumns().getByIndex(2).OptimalWidth = True
	
	end if

End Sub
