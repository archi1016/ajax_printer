# ajax_printer
ajax_printer

Use AJAX for browser to print receipt

![doc_1](doc-1.png)


Query url

	GET http://127.0.0.1:35427/?q=list_printer
		RETURN Printers Information
		
	POST http://127.0.0.1:35427/?q=preview
		POST Xml struct
		RETURN png to display (only last page)
		
	POST http://127.0.0.1:35427/?q=print
		POST xml struct


XML struct

	<paper printer="{PRINTER}" copies="{COPIES}">
		<l c="set_font">Arial</l>
		<l c="set_size">6</l>
		
		<l c="padding_top">{PADDING_TOP}</l>
		<l c="padding_left">{PADDING_LEFT}</l>
		<l c="padding_right">{PADDING_RIGHT}</l>
		
		<l c="print" size="10">{NAME}</l>
		<l c="scroll_down">2</l>
		<l c="line">1</l>
		
		<l c="print_vt" x="84%" y="20%" align="left" font="Free 3 of 9" size="19">*{BARCODE}*</l>
		<l c="print_vt" x="94%" y="65%" align="left" font="Consolas" size="10">{BARCODE}</l>
		<l c="print_vt" x="94%" y="24%" align="left" font="Consolas" size="10">{PRICE}</l>
		
		<l c="scroll_down">3</l>
		<l c="fields">
			<f c="print" w="20%">111</f>
			<f c="print" w="60%">aaaaa</f>
		</l>
		<l c="fields">
			<f c="print" w="20%">222</f>
			<f c="print" w="60%">bbbbb</f>
		</l>
		<l c="fields">
			<f c="print" w="20%">333</f>
			<f c="print" w="60%">ccccc</f>
		</l>
		<l c="fields">
			<f c="print" w="20%">444</f>
			<f c="print" w="60%">ddddd</f>
		</l>
		
		<l c="print_xy" x="87%" y="3%" border="0.5">臺灣製</l>
		
		<l c="absolute_y">72%</l>
		<l c="line_w" begin="0%" end="80%">0.5</l>
		<l c="scroll_down">1</l>
		<l c="fields">
			<f c="print" w="20%">111</f>
			<f c="print" w="60%">aaaaa</f>
		</l>
		<l c="fields">
			<f c="print" w="20%">222</f>
			<f c="print" w="60%">bbbbb</f>
		</l>
		<l c="fields">
			<f c="print" w="20%">333</f>
			<f c="print" w="60%">ccccc</f>
		</l>
		<l c="fields">
			<f c="print" w="20%">444</f>
			<f c="print" w="60%">ddddd</f>
		</l>

		<l c="done" />
	</paper>


Command

	new_page
		換頁

		
	done
		列印

		
	padding_top
		上邊界留白
		例: 2 (2pt)
		例: 5% (5% 的頁寬)

		
	padding_bottom
		下邊界留白
		例: 2 (2pt)
		例: 5% (5% 的頁寬)
		
		
	padding_left
		左邊界留白
		例: 2 (2pt)
		例: 5% (5% 的頁寬)
		

	padding_right
		右邊界留白
		例: 2 (2pt)
		例: 5% (5% 的頁寬)

	
		set_font
		設定預設字體名稱
		例: Arial

		
	push_font
		預設字體名稱推入堆疊

	
	pop_font
		從堆疊取回預設字體名稱
		
		
	set_size
		設定預設字體大小
		例: 16 (16pt)

		
	push_size
		預設字體大小推入堆疊

		
	pop_size
		從堆疊取回預設字體大小

		
	push_xy
		XY 座標推入堆疊

		
	pop_xy
		從堆疊取回 XY 座標

		
	absolute_x
		設定 X 座標 (整頁)
		例: 2 (2pt)
		例: 5% (5% 的頁寬)

		
	absolute_y
		設定 Y 座標 (整頁)
		例: 4 (4pt)
		例: 9% (9% 的頁寬)

		
	relative_x
		設定 X 座標 (以目前的 X 座標 )
		例: 2 (2pt)
		例: 5% (5% 的頁寬)

		
	relative_y
		設定 Y 座標 (以目前的 Y 座標 )
		例: 4 (4pt)
		例: 9% (9% 的頁寬)

		
	scroll_up
		同 relative_y (自動換算成負值)

		
	scroll_down
		同 relative_y


	line
		例: 1 (1pt)
		例: 0.1% (0.1% 的頁寬)


	line_w
		例: 1 (1pt)
		例: 0.1% (0.1% 的頁寬)
		
			begin
				開始位置
				例: 1 (1pt)
				例: 10% (10% 的頁寬)
			
			end
				結束位置
				例: 100 (100pt)
				例: 80% (80% 的頁寬)
		
		
	print
		列印文字, 可多行
		
			font (非必要)
				指定字體
				
			size (非必要)
				指定文字大小
				例: 16 (16pt)
				
			border (非必要)
				邊框
				例: 2 (2pt)


	print_i
		反白列印文字, 可多行
		
			font (非必要)
				指定字體
				
			size (非必要)
				指定文字大小
				例: 16 (16pt)
				
			border (非必要)
				邊框
				例: 2 (2pt)


	print_center
		置中列印文字, 可多行

			font (非必要)
				指定字體
				
			size (非必要)
				指定文字大小
				例: 16 (16pt)
				
			border (非必要)
				邊框
				例: 2 (2pt)
				

	print_center_i
		置中反白列印文字, 可多行

			font (非必要)
				指定字體
				
			size (非必要)
				指定文字大小
				例: 16 (16pt)
				
			border (非必要)
				邊框
				例: 2 (2pt)

		
	print_right
		靠右列印文字, 可多行

			font (非必要)
				指定字體
				
			size (非必要)
				指定文字大小
				例: 16 (16pt)
				
			border (非必要)
				邊框
				例: 2 (2pt)

				
	print_right_i
		靠右反白列印文字, 可多行

			font (非必要)
				指定字體
				
			size (非必要)
				指定文字大小
				例: 16 (16pt)
				
			border (非必要)
				邊框
				例: 2 (2pt)

				
	print_xy
		指定位置列印文字, 可多行
		
			x
				指定 X 座標 (整頁)
				例: 2 (2pt)
				例: 5% (5% 的頁寬)

			y
				指定 Y 座標 (整頁)
				例: 5 (5pt)
				例: 8% (8% 的頁高)

			font (非必要)
				指定字體
				
			size (非必要)
				指定文字大小
				例: 16 (16pt)
				
			border (非必要)
				邊框
				例: 2 (2pt)
			
	print_vt
		指定位置列印旋轉文字, 可多行
	
			x
				指定 X 座標 (整頁)
				例: 2 (2pt)
				例: 5% (5% 的頁寬)

			y
				指定 Y 座標 (整頁)
				例: 5 (5pt)
				例: 8% (8% 的頁高)
				
			align
				旋轉方式
					LEFT (逆時鐘轉 90 度)
					RIGHT  (順時鐘轉 90 度)
					BOTTOM (轉 180 度)

			font (非必要)
				指定字體
				
			size (非必要)
				指定文字大小
				例: 16 (16pt)
				
			border (非必要)
				邊框
				例: 2 (2pt)
			
 
	barcode
		待修正

		
	barcode_center
		待修正

		
	barcode_right
		待修正

		
	fields
		待補
		
