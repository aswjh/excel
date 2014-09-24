#excel for golang

read and write excel for golang.

go语言读写excel.

#dependency

[github.com/mattn/go-ole][ole]

#install

go get github.com/aswjh/excel

#example
``` go
package main

import (
	"time"
	"github.com/aswjh/excel"
)

func main() {
	option := excel.Option{Visible: true, ScreenUpdating: true, DisplayAlerts: true}
	xl, _ := excel.New(option)      //xl := excel.Open("test_excel.xls", option)
	defer xl.Quit()

	sheet, _ := xl.Sheet(1)         //xl.Sheet("sheet1")
	for i:=2; i<=6; i++ {
		sheet.Cells(2, i, 1000+i)
	}

	println("cell strings: ", sheet.Cells(2, 1), sheet.Cells(2, 2), sheet.Cells(2, 3))
	time.Sleep(3000000000)

	xl.SaveAs("test_excel.xls")
}

```

#license

The [BSD 3-Clause license][bsd]

[ole]: http://github.com/mattn/go-ole
[bsd]: http://opensource.org/licenses/BSD-3-Clause

