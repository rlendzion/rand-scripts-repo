Dim avg_x,avg_y,sum_x_delta_sqrd,sum_product,slope,intercept,result,arr
Dim a(19,1)
for i = 0 to 19
	for j = 0 to 1
		a(i,j) = i+1+((j*i)*j*3) 
	Next
Next
for d = 0 to 19
	for w = 0 to 1
		arr = arr & a(d,w) & vbTab 'tab
	Next
	arr = arr & vbCrLf 'new line
Next
'Print array
MsgBox arr
for i = 0 to 19
	avg_x = avg_x + a(i,0)
Next
avg_x = avg_x / 20
for i = 0 to 19
	avg_y = avg_y + a(i,1)
Next
avg_y = avg_y / 20
MsgBox "Average X: " & avg_x & "; Average Y: " & avg_y
for i = 0 to 19
	sum_product = sum_product + (a(i,0)-avg_x) * (a(i,1)-avg_y)
Next
MsgBox "The sum of product of deviations: " & sum_product
for i = 0 to 19
	sum_x_delta_sqrd = sum_x_delta_sqrd + (a(i,0)-avg_x)^2
Next
MsgBox "Squared deviation from the mean (X): " & sum_x_delta_sqrd
MsgBox "Now that all is ready, let's calculate y' for x=21"
slope = sum_product/sum_x_delta_sqrd
intercept = avg_y - slope*avg_x
MsgBox "Y-intercept: " & intercept & vbCrLf & "Slope: " & slope
result = slope*21+intercept
MsgBox "Predicted Y value for X=21 is: " & result
