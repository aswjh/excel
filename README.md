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
	option := excel.Option{"Visible": true, "DisplayAlerts": true}
	xl, _ := excel.New(option)      //xl := excel.Open("test_excel.xls", option)
	defer xl.Quit()

	sheet, _ := xl.Sheet(1)         //xl.Sheet("sheet1")
	for i:=2; i<=6; i++ {
		sheet.MustCells(2, i, 1000+i)
	}

	cell := sheet.MustCell(5, 6)
    cell.Put("go")
	cell.Put("font", map[string]interface{}{"name": "Times New Roman", "size": "26", "bold": true})
	cell.Put("interior", map[string]interface{}{"colorindex": 6})

	println("cell strings:", sheet.MustCells(2, 2), sheet.MustCells(2, 3))
	time.Sleep(3000000000)

	xl.SaveAs("test_excel.xls")
}

```

#license

The [BSD 3-Clause license][bsd]

[ole]: http://github.com/mattn/go-ole
[bsd]: http://opensource.org/licenses/BSD-3-Clause


