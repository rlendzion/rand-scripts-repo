''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Convolution logic in vbScript using standard kernel (1,0,-1)
'Note: the output table does not show any edges, because the
'values of a matrix form some sort of a gradient (can be chan-
'ged with different function inputting values into the matrix
'Author: Robert Lendzion
'Date: 2019-11-28
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim nInput
nInput = InputBox("Define the size of a 2D matrix (at least 6):",,"6")-1
'Create 6x6 matrix
ReDim n(nInput,nInput) 
for r = 0 to ubound(n)
	for c = 0 to ubound(n)
	n(r,c) = r*c
	Next
Next
MsgBox "Matrix created successfully"

'Print the matrix
Dim matrix
for lngth = 0 to UBound(n)
	for depth = 0 to UBound(n)
		matrix = matrix & n(lngth,depth)  & vbTab
	Next
	matrix = matrix & vbCrLf
Next
MsgBox matrix

'Create 3x3 kernel (filter)
ReDim k(2,2)
for i = 0 to ubound(k)
	for c = 0 to ubound(k)
		k(i,c) = (c-1)*-1
	Next
Next

'Print first row of a kernel for check
Dim t 
t = " | "
for i = 0 to ubound(k)
	t = t & cStr(k(0,i)) & " | "
Next
MsgBox "First row of a kernel: " & t


Dim output_shape 
output_shape = cInt(ubound(n))-cInt(ubound(k))+1 'define the shape of an output
MsgBox "Shape of output: " & output_shape & " x " & output_shape

'Convolve with forward propagation
ReDim output(output_shape,output_shape)
Dim calc 'value that will be put to the final table
for go_down = 0 to output_shape-1
	for go_right = 0 to output_shape-1 'move one step to the right
		calc = 0
		for col = 0 to ubound(k)
			for row = 0 to ubound(k) 'limit of kernel
				calc = calc + n(row+go_down,col+go_right) * k(row,col)  'sum up all element-wise products for every cell of the output table
			Next
		Next
	output(go_down,go_right) = calc
	'MsgBox "value obtained for output: " & calc & "; going in "& go_down+1 & "th row, " & go_right+1 & "th column." 'hide this if you don't want to see any item being added to the output
	Next
Next

'Convoluted output
Dim conv
for lngth = 0 to UBound(output)
	for depth = 0 to UBound(output)
		conv = conv & output(lngth,depth)  & vbTab
	Next
	conv = conv & vbCrLf
Next
MsgBox conv
