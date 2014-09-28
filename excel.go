package excel

import (
    "strconv"
    "strings"
    "path/filepath"
    "unsafe"
    "reflect"
    "errors"
    "fmt"
    "github.com/mattn/go-ole"
    "github.com/mattn/go-ole/oleutil"
)

type Option map[string]interface{}

type MSO struct {
    Option
    IuApp                      *ole.IUnknown
    IdExcel                    *ole.IDispatch
    IdWorkBooks                *ole.IDispatch
    Version                   float64
    FILEFORMAT          map[string]int
}

type WorkBook struct {
    Idisp *ole.IDispatch
    *MSO
}

type WorkBooks []WorkBook

type Sheet struct {
    Idisp *ole.IDispatch
}

type VARIANT struct {
    *ole.VARIANT
}

//get val of MS VARIANT
func (va VARIANT) Value() (val interface{}) {
    switch va.VT {
        case 2:
            val = *((*int16)(unsafe.Pointer(&va.Val)))
        case 3:
            val = *((*int32)(unsafe.Pointer(&va.Val)))
        case 4:
            val = *((*float32)(unsafe.Pointer(&va.Val)))
        case 5:
            val =*((*float64)(unsafe.Pointer(&va.Val)))
        case 8:                     //string
            val = *((**uint16)(unsafe.Pointer(&va.Val)))
        case 11:
            val = *((*bool)(unsafe.Pointer(&va.Val)))
        case 16:
            val = *((*int8)(unsafe.Pointer(&va.Val)))
        case 17:
            val = *((*uint8)(unsafe.Pointer(&va.Val)))
        case 18:
            val = *((*uint16)(unsafe.Pointer(&va.Val)))
        case 19:
            val = *((*uint32)(unsafe.Pointer(&va.Val)))
        case 20:
            val = *((*int64)(unsafe.Pointer(&va.Val)))
        case 21:
            val = *((*uint64)(unsafe.Pointer(&va.Val)))
    }
    return
}

//
func String(val interface{}) (ret string) {
    switch val.(type) {
        case int16:
            ret = strconv.FormatInt(int64(val.(int16)), 10)
        case int32:
            ret = strconv.FormatInt(int64(val.(int32)), 10)
        case float32:
            ret = strconv.FormatFloat(float64(val.(float32)), 'f', 2, 64)
        case float64:
            ret = strconv.FormatFloat(val.(float64), 'f', 2, 64)
        case *uint16:                     //string
            ret = ole.UTF16PtrToString(val.(*uint16))
        case bool:
            if val.(bool) {
                ret = "true"
            } else {
                ret = "false"
            }
        case int8:
            ret = strconv.FormatInt(int64(val.(int8)), 10)
        case uint8:
            ret = strconv.FormatUint(uint64(val.(uint8)), 10)
        case uint16:
            ret = strconv.FormatUint(uint64(val.(uint16)), 10)
        case uint32:
            ret = strconv.FormatUint(uint64(val.(uint32)), 10)
        case int64:
            ret = strconv.FormatInt(val.(int64), 10)
        case uint64:
            ret = strconv.FormatUint(val.(uint64), 10)
    }
    return
}

//
func Init(options... Option) (mso *MSO) {
    ole.CoInitialize(0)
    app, _ := oleutil.CreateObject("Excel.Application")
    excel, _ := app.QueryInterface(ole.IID_IDispatch)
    wbs := oleutil.MustGetProperty(excel, "WorkBooks").ToIDispatch()
    ver, _ := strconv.ParseFloat(oleutil.MustGetProperty(excel, "Version").ToString(), 64)

    option := Option{"Visible": true, "DisplayAlerts": true, "ScreenUpdating": true}
    if options != nil {
        option = options[0]
    }
    mso = &MSO{Option:option, IuApp:app, IdExcel:excel, IdWorkBooks:wbs, Version:ver}
    mso.SetOption(option, 1)

    //XlFileFormat Enumeration: http://msdn.microsoft.com/en-us/library/office/ff198017%28v=office.15%29.aspx
    mso.FILEFORMAT = map[string]int {"txt":-4158, "csv":6, "html":44, "xlsx":51, "xls":56}
    return
}

//
func New(options... Option) (mso *MSO, err error){
    defer Except("New", &err)
    mso = Init(options...)
    _, err = mso.AddWorkBook()
    return
}

//
func Open(full string, options... Option) (mso *MSO, err error) {
    defer Except("Open", &err)
    mso = Init(options...)
    _, err = mso.OpenWorkBook(full)
    return
}

//
func (mso *MSO) Save() ([]error) {
    return mso.WorkBooks().Save()
}

//
func (mso *MSO) SaveAs(args... interface{}) ([]error) {
    return mso.WorkBooks().SaveAs(args...)
}

//
func (mso *MSO) Quit() (err error) {
    defer Except("Quit", &err, ole.CoUninitialize)
    if r := recover(); r != nil {   //catch panic of which defering Quit.
        info := fmt.Sprintf("***panic before Quit: %+v", r)
        fmt.Println(info)
        err = errors.New(info)
    }
    oleutil.MustCallMethod(mso.IdWorkBooks, "Close")
    oleutil.MustCallMethod(mso.IdExcel, "Quit")
    mso.IdWorkBooks.Release()
    mso.IdExcel.Release()
    mso.IuApp.Release()
    return
}

//
func (mso *MSO) SetOption(args... interface{}) (err error) {
    defer Except("SetOption", &err)
    leng := len(args)
    if leng>0 {
        opts, curs, isinit := mso.Option, Option{}, false
        if options, ok := args[0].(Option); ok {
            curs, isinit = options, leng==2 && args[1].(int) > 0
        } else if key, ok := args[0].(string); ok {
            curs[key] = args[1]
        }
        for key, val := range curs {
            if isinit {
                oleutil.PutProperty(mso.IdExcel, key, val)
            } else if opt, ok := opts[key]; ! ok || val != opt {
                oleutil.PutProperty(mso.IdExcel, key, val)
                opts[key] = val
            }
        }
    }
    return
}

//
func (mso *MSO) Pick(workx string, id interface{}) (ret *ole.IDispatch, err error) {
    defer Except("Pick", &err)
    if id_int, ok := id.(int); ok {
        ret = oleutil.MustGetProperty(mso.IdExcel, workx, id_int).ToIDispatch()
    } else if id_str, ok := id.(string); ok {
        ret = oleutil.MustGetProperty(mso.IdExcel, workx, id_str).ToIDispatch()
    } else {
        err = errors.New("sheet id incorrect")
    }
    return
}

//
func (mso *MSO) CountWorkBooks() (int) {
    return (int)(oleutil.MustGetProperty(mso.IdWorkBooks, "Count").Val)
}

//
func (mso *MSO) WorkBooks() (wbs WorkBooks) {
    num := mso.CountWorkBooks()
    for i:=1; i<=num; i++ {
        wbs = append(wbs, WorkBook{oleutil.MustGetProperty(mso.IdExcel, "WorkBooks", i).ToIDispatch(), mso})
    }
    return
}

//
func (mso *MSO) AddWorkBook() (WorkBook, error) {
    _wb , err := oleutil.CallMethod(mso.IdWorkBooks, "Add")
    defer Except("AddWorkBook", &err)
    return WorkBook{_wb.ToIDispatch(), mso}, err
}

//
func (mso *MSO) OpenWorkBook(full string) (WorkBook, error) {
    _wb, err := oleutil.CallMethod(mso.IdWorkBooks, "open", full)
    defer Except("OpenWorkBook", &err)
    return WorkBook{_wb.ToIDispatch(), mso}, err
}

//
func (mso *MSO) ActivateWorkBook(id interface{}) (wb WorkBook, err error) {
    defer Except("ActivateWorkBook", &err)
    _wb, _ := mso.Pick("WorkBooks", id)
    wb = WorkBook{_wb, mso}
    wb.Activate()
    return
}

//
func (mso *MSO) ActiveWorkBook() (WorkBook, error) {
    _wb, err := oleutil.GetProperty(mso.IdExcel, "ActiveWorkBook")
    return WorkBook{_wb.ToIDispatch(), mso}, err
}

//
func (mso *MSO) CountSheets() (int) {
    sheets := oleutil.MustGetProperty(mso.IdExcel, "Sheets").ToIDispatch()
    return (int)(oleutil.MustGetProperty(sheets, "Count").Val)
}

//
func (mso *MSO) Sheets() (sheets []Sheet) {
    num := mso.CountSheets()
    for i:=1; i<=num; i++ {
        sheet := Sheet{oleutil.MustGetProperty(mso.IdExcel, "WorkSheets", i).ToIDispatch()}
        sheets = append(sheets, sheet)
    }
    return
}

//
func (mso *MSO) Sheet(id interface {}) (Sheet, error) {
    _sheet, err := mso.Pick("WorkSheets", id)
    return Sheet{_sheet}, err
}

//
func (mso *MSO) AddSheet(wb *ole.IDispatch, args... string) (Sheet, error) {
    sheets := oleutil.MustGetProperty(wb, "Sheets").ToIDispatch()
    defer sheets.Release()
    _sheet, err := oleutil.CallMethod(sheets, "Add")
    sheet := Sheet{_sheet.ToIDispatch()}
    if args != nil {
        sheet.Name(args...)
    }
    sheet.Select()
    return sheet, err
}

//
func (mso *MSO) SelectSheet(id interface{}) (sheet Sheet, err error) {
    sheet, err = mso.Sheet(id)
    sheet.Select()
    return
}

//
func (wbs WorkBooks) Save() (errs []error) {
    for _, wb := range wbs {
        errs = append(errs, wb.Save())
    }
    return
}

//
func (wbs WorkBooks) SaveAs(args... interface{}) (errs []error) {
    num := len(wbs)
    if num<2 {
        for _, wb := range wbs {
            errs = append(errs, wb.SaveAs(args...))
        }
    }  else {
        full := args[0].(string)
        ext := filepath.Ext(full)
        for i, wb := range wbs {
            args[0] = strings.Replace(full, ext, "_"+strconv.Itoa(i)+ext, 1)
            errs = append(errs, wb.SaveAs(args...))
        }
    }
    return
}

//
func (wbs WorkBooks) Close() (errs []error) {
    for _, wb := range wbs {
        errs = append(errs, wb.Close())
    }
    return
}

//
func (wb WorkBook) Activate() (err error) {
    defer Except("WorkBook.Activate", &err)
    _, err = oleutil.CallMethod(wb.Idisp, "Activate")
    return
}

//
func (wb WorkBook) Name() (string) {
    defer NoExcept()
    return oleutil.MustGetProperty(wb.Idisp, "Name").ToString()
}

//
func (wb WorkBook) Save() (err error) {
    defer Except("WorkBook.Save", &err)
    _, err = oleutil.CallMethod(wb.Idisp, "Save")
    return
}

//
func (wb WorkBook) SaveAs(args... interface{}) (err error) {
    defer Except("WorkBook.SaveAs", &err)
    _, err = oleutil.CallMethod(wb.Idisp, "SaveAs", args...)
    return
}

//
func (wb WorkBook) Close() (err error) {
    defer Except("WorkBook.Close", &err)
    _, err = oleutil.CallMethod(wb.Idisp, "Close")
    return
}

//
func (sheet Sheet) Select() (err error) {
    defer Except("Sheet.Select", &err)
    _, err = oleutil.CallMethod(sheet.Idisp, "Select")
    return
}

//
func (sheet Sheet) Delete() (err error) {
    defer Except("Sheet.Delete", &err)
    _, err = oleutil.CallMethod(sheet.Idisp, "Delete")
    return
}

//
func (sheet Sheet) Name(args... string) (name string) {
    defer NoExcept()
    if args == nil {
        name = oleutil.MustGetProperty(sheet.Idisp, "Name").ToString()
    } else {
        name = args[0]
        oleutil.MustPutProperty(sheet.Idisp, "Name", name)
    }
    return
}

//get Property as interface.
func (sheet Sheet) GetCells(r int, c int, args... string) (ret interface{} , err error) {
    defer Except("Sheet.GetCells", &err)
    cell := oleutil.MustGetProperty(sheet.Idisp, "Range", Cell2r(r, c)).ToIDispatch()
    defer NoExcept(cell.Release)
    if args == nil {
        ret = VARIANT{oleutil.MustGetProperty(cell, "Value")}.Value()
    } else {
        ret = VARIANT{oleutil.MustGetProperty(cell, args[0])}.Value()
    }
    return
}

//Must get Property as interface.
func (sheet Sheet) MustGetCells(r int, c int, args... string) (ret interface{}) {
    ret, err := sheet.GetCells(r, c, args...)
    if err != nil {
        panic(err.Error())
    }
    return
}

//put Property.
func (sheet Sheet) PutCells(r int, c int, args... interface{}) (err error) {
    defer Except("Sheet.PutCells", &err)
    cell := oleutil.MustGetProperty(sheet.Idisp, "Range", Cell2r(r, c)).ToIDispatch()
    defer NoExcept(cell.Release)
    num := len(args)
    if num == 1 {
        oleutil.MustPutProperty(cell, "Value", args[0])
    } else if num > 1 {
        oleutil.MustPutProperty(cell, args[0].(string), args[1])
    }
    return
}

//get value as string/put value.
func (sheet Sheet) Cells(r int, c int, vals... interface{}) (ret string, err error) {
    defer Except("Sheet.Cells", &err)
    if vals == nil {
        v := sheet.MustGetCells(r, c)
        if v != nil {
            ret = String(v)
        }
    } else {
        err = sheet.PutCells(r, c, vals[0])
    }
    return
}

//Must get value as string/Must put value.
func (sheet Sheet) MustCells(r int, c int, vals... interface{}) (ret string) {
    ret, err := sheet.Cells(r, c, vals...)
    if err != nil {
        panic(err.Error())
    }
    return
}

//convert (5,27) to "AA5", xp+office2003 cant use cell(1,1).
func Cell2r(x, y int) (ret string) {
    for y>0 {
        a, b := int(y/26), y%26
        if b==0 {
            a-=1
            b=26
        }
        ret = string(rune(b+64))+ret
        y = a
    }
    return ret+strconv.Itoa(x)
}

//
func Except(info string, err *error, funcs... interface{}) {
    r := recover()
    if err != nil {
        if r != nil {
            *err = errors.New(fmt.Sprintf("@"+info+": %+v", r))
        } else if *err != nil {
            *err = errors.New("%"+info+"%"+(*err).Error())
        }
    }
    if funcs != nil {
        NoExcept(funcs...)
    }
}

//
func NoExcept(funcs... interface{}) {
    recover()
    if funcs != nil {
        prei, args := -1, []reflect.Value{}
        for i, one := range funcs {
            cur := reflect.ValueOf(one)
            if cur.Kind().String() == "func" {
                if prei > -1 {
                    RftCall(reflect.ValueOf(funcs[prei]), args...)
                }
                prei, args = i, []reflect.Value{}
            } else {
                args = append(args, cur)
            }
        }
        if prei > -1 {
            RftCall(reflect.ValueOf(funcs[prei]), args...)
        }
    }
}

//
func RftCall(function reflect.Value, args... reflect.Value) (err error) {
    defer func() {
        if r := recover(); r != nil {
            err = errors.New(fmt.Sprintf("RftCall panic: %+v", r))
        }
    }()
    function.Call(args)
    return
}





