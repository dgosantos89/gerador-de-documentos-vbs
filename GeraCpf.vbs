Function geraCpf(comPontos)
	Dim max
	Dim min
	Dim n1, n2, n3, n4, n5, n6, n7, n8, n9
	Dim d1, d2
	
	' Define o range
	max = 9
	min = 0

	Randomize ' Inicializa o gerador de numeros randomicos
	n1 = Int((max-min+1)*rnd+min)
	n2 = Int((max-min+1)*rnd+min)
	n3 = Int((max-min+1)*rnd+min)
	n4 = Int((max-min+1)*rnd+min)
	n5 = Int((max-min+1)*rnd+min)
	n6 = Int((max-min+1)*rnd+min)
	n7 = Int((max-min+1)*rnd+min)
	n8 = Int((max-min+1)*rnd+min)
	n9 = Int((max-min+1)*rnd+min)

	' Calcula o primeiro digito validador
	d1 = (n9*2)+(n8*3)+(n7*4)+(n6*5)+(n5*6)+(n4*7)+(n3*8)+(n2*9)+(n1*10)
	d1 = 11 - ( d1 Mod 11 )
	
	If d1 >= 10 Then
		d1 = 0
	End If

	' Calcula o segundo digito validador
	d2 = (d1*2)+(n9*3)+(n8*4)+(n7*5)+(n6*6)+(n5*7)+(n4*8)+(n3*9)+(n2*10)+(n1*11)
	d2 = 11 - ( d2 mod 11 )

	If d2 >= 10 Then
		d2 = 0
	End If
 
 	If comPontos Then
		' Retorna CPF com pontos
		geraCpf = ""&n1&n2&n3&"."&n4&n5&n6&"."&n7&n8&n9&"."&n10&n11&n12&"-"&d1&d2
	Else
	 	' Retorna CPF sem pontos
		geraCpf = ""&n1&n2&n3&n4&n5&n6&n7&n8&n9&n10&n11&n12&d1&d2
	End If
End Function