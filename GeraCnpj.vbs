Function geraCnpj(comPontos)
	Dim max
	Dim min
	Dim n1, n2, n3, n4, n5, n6, n7, n8, n9, n10, n11, n12
	Dim d1, d2
	
	' Define o range
	max = 9
	min = 0

	Randomize
	n1 = Int((max-min+1)*rnd+min)
	n2 = Int((max-min+1)*rnd+min)
	n3 = Int((max-min+1)*rnd+min)
	n4 = Int((max-min+1)*rnd+min)
	n5 = Int((max-min+1)*rnd+min)
	n6 = Int((max-min+1)*rnd+min)
	n7 = Int((max-min+1)*rnd+min)
	n8 = Int((max-min+1)*rnd+min)
	n9 = 0
	n10 = 0
	n11 = 0
	n12 = 0

	' Calcula o primeiro digito validador
	d1 = (n12*2)+(n11*3)+(n10*4)+(n9*5)+(n8*6)+(n7*7)+(n6*8)+(n5*9)+(n4*2)+(n3*3)+(n2*4)+(n1*5)
	d1 = 11 - ( d1 mod 11 )
	If d1 >= 10 Then
		d1 = 0
	End If
	
	' Calcula o segundo digito validador
	d2 = (d1*2)+(n12*3)+(n11*4)+(n10*5)+(n9*6)+(n8*7)+(n7*8)+(n6*9)+(n5*2)+(n4*3)+(n3*4)+(n2*5)+(n1*6)
	d2 = 11 - ( d2 mod 11 )
	if d2 >= 10 Then
		d2 = 0
	End If

	If comPontos Then
		' Retorna CNPJ com pontos
		geraCnpj = ""&n1&n2&"."&n3&n4&n5&"."&n6&n7&n8&"/"&n9&n10&n11&n12&"-"&d1&d2
	Else
		' Retorna CNPJ sem pontos
		geraCnpj = ""&n1&n2&n3&n4&n5&n6&n7&n8&n9&n10&n11&n12&d1&d2
	End If
End Function