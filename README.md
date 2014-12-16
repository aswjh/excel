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
	println("cell strings:", sheet.MustCells(2, 2), sheet.MustCells(2, 3))

	cell := sheet.MustCell(5, 6)
    cell.Put("go")
	cell.Put("font", map[string]interface{}{"name": "Arial", "size": 26, "bold": true})
	cell.Put("interior", "colorindex", 6)

	time.Sleep(3000000000)
	xl.SaveAs("test_excel.xls")      //xl.SaveAs("test_excel", "html")
}

```

#license

The [BSD 3-Clause license][bsd]

[ole]: http://github.com/mattn/go-ole
[bsd]: http://opensource.org/licenses/BSD-3-Clause




