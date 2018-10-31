package excel

import (
    "os"
    "syscall"
    "time"
    "strconv"
    "strings"
    "path/filepath"
    "unsafe"
    "reflect"
    "errors"
    "fmt"
    "runtime/debug"
    "github.com/go-ole/go-ole"
    "github.com/go-ole/go-ole/oleutil"
)

var (
    modoleaut32, _ = syscall.LoadDLL("oleaut32.dll")
    procSafeArrayGetElement, _ = modoleaut32.FindProc("SafeArrayGetElement")
    procSafeArrayGetVartype, _ = modoleaut32.FindProc("SafeArrayGetVartype")
)

type Option map[string]interface{}

type MSO struct {
    Option
    IuApp                      *ole.IUnknown
    IdExcel                    *ole.IDispatch
    IdWorkBooks                *ole.IDispatch
    WorkBook WorkBook
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

type Range struct {
    Idisp *ole.IDispatch
}

type Cell struct {
    Idisp *ole.IDispatch
}

type VARIANT struct {
    *ole.VARIANT
}

//
func Initialize(opt... Option) (mso *MSO) {
    ole.CoInitialize(0)
    app, _ := oleutil.CreateObject("Excel.Application")
    excel, _ := app.QueryInterface(ole.IID_IDispatch)
    wbs := oleutil.MustGetProperty(excel, "WorkBooks").ToIDispatch()
    ver, _ := strconv.ParseFloat(oleutil.MustGetProperty(excel, "Version").ToString(), 64)

    if len(opt) == 0 {
        opt = []Option {{"Visible": true, "DisplayAlerts": true, "ScreenUpdating": true}}
    }
    mso = &MSO{Option:opt[0], IuApp:app, IdExcel:excel, IdWorkBooks:wbs, Version:ver}
    mso.SetOption(1)

    //XlFileFormat Enumeration: http://msdn.microsoft.com/en-us/library/office/ff198017%28v=office.15%29.aspx
    mso.FILEFORMAT = map[string]int {"txt":-4158, "csv":6, "html":44}
    return
}

//
func New(opt... Option) (mso *MSO, err error){
    defer Except("New", &err)
    mso = Initialize(opt...)
    mso.WorkBook, err = mso.AddWorkBook()
    return
}

//
func Open(full string, opt... Option) (mso *MSO, err error) {
    defer Except("Open", &err)
    mso = Initialize(opt...)
    mso.WorkBook, err = mso.OpenWorkBook(full)
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
        err = errors.New(fmt.Sprintf("***panic before Quit: %+v", r))
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
    ops, curs, isinit := mso.Option, Option{}, false
    leng := len(args)
    if leng==1 {
        switch args[0].(type) {
            case int:
                if args[0].(int)>0 {
                    curs, isinit = ops, true
                }
            case Option:
                curs = args[0].(Option)
        }
    } else if leng==2 {
        if key, ok := args[0].(string); ok {
            curs[key] = args[1]
        }
    }
    for key, val := range curs {
        if isinit {
            _, err = oleutil.PutProperty(mso.IdExcel, key, val)
        } else if one, ok := ops[key]; ! ok || val != one {
            _, err = oleutil.PutProperty(mso.IdExcel, key, val)
            ops[key] = val
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
    defer Except("", nil)
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
    if len(args)>1 {
        switch args[1].(type) {
            case string:
                ffs := strings.ToLower(args[1].(string))
                if n, ok := wb.FILEFORMAT[ffs]; ok {
                    fn := args[0].(string)
                    if ! strings.HasSuffix(fn, "."+ffs) {
                        args[0] = fn+"."+ffs
                    }
                    args[1] = n
                } else {
                    args[1] = nil
                }
        }
    }
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
func (sheet Sheet) Release() {
    sheet.Idisp.Release()
}

//
func (sheet Sheet) Name(args... string) (name string) {
    defer Except("", nil)
    if len(args) == 0 {
        name = oleutil.MustGetProperty(sheet.Idisp, "Name").ToString()
    } else {
        name = args[0]
        oleutil.MustPutProperty(sheet.Idisp, "Name", name)
    }
    return
}

//get cell Property as interface.
func (sheet Sheet) GetCell(r int, c int, args... string) (ret interface{} , err error) {
    defer Except("Sheet.GetCell", &err)
    cell := sheet.MustCell(r, c)
    defer DoFuncs(cell.Release)
    ret, err = cell.Get(args...)
    return
}

//Must get cell Property as interface.
func (sheet Sheet) MustGetCell(r int, c int, args... string) (interface{}) {
    ret, err := sheet.GetCell(r, c, args...)
    if err != nil {
        panic(err)
    }
    return ret
}

//put cell Property.
func (sheet Sheet) PutCell(r int, c int, args... interface{}) (err error) {
    defer Except("Sheet.PutCell", &err)
    cell := sheet.MustCell(r, c)
    defer DoFuncs(cell.Release)
    err = cell.Put(args...)
    return
}

//get cell Property as string, put cell Property.
func (sheet Sheet) Cells(r int, c int, vals... interface{}) (ret string, err error) {
    defer Except("Sheet.Cells", &err)
    if len(vals) == 0 {
        ret = String(sheet.MustGetCell(r, c))
    } else {
        err = sheet.PutCell(r, c, vals[0])
    }
    return
}

//Must get cell Property as string, Must put cell Property.
func (sheet Sheet) MustCells(r int, c int, vals... interface{}) (ret string) {
    ret, err := sheet.Cells(r, c, vals...)
    if err != nil {
        panic(err)
    }
    return
}

//get cell pointer.
func (sheet Sheet) Cell(r int, c int) (cell Cell, err error) {
    defer Except("Sheet.Cell", &err)
    _cell, err := oleutil.GetProperty(sheet.Idisp, "Cells", r, c)
    cell = Cell{_cell.ToIDispatch()}
    return
}

//Must get cell pointer.
func (sheet Sheet) MustCell(r int, c int) (cell Cell) {
    cell = Cell{oleutil.MustGetProperty(sheet.Idisp, "Cells", r, c).ToIDispatch()}
    return
}

//get range pointer.
func (sheet Sheet) Range(rang string) (Range) {
    return Range{oleutil.MustGetProperty(sheet.Idisp, "Range", rang).ToIDispatch()}
}

//get range Property as interface.
func (sheet Sheet) GetRange(rang string) (ret interface{} , err error) {
    defer Except("Sheet.GetRange", &err)
    rg := sheet.Range(rang)
    defer DoFuncs(rg.Release)
    ret, err = rg.Get()
    return
}

func (sheet Sheet) MustGetRange(rang string) (interface{}) {
    ret, err := sheet.GetRange(rang)
    if err != nil {
        panic(err)
    }
    return ret
}

//put range Property.
func (sheet Sheet) PutRange(rang string, args... interface{}) (err error) {
    defer Except("Sheet.PutRange", &err)
    rg := sheet.Range(rang)
    defer DoFuncs(rg.Release)
    err = rg.Put(args...)
    return
}

//get sheet Property
func (sheet Sheet) Get(args... string) (ret interface{}, err error) {
    ret, err = GetProperty(sheet.Idisp, args...)
    return
}

//Must get sheet Property
func (sheet Sheet) MustGet(args... string) (interface{}) {
    return MustGetProperty(sheet.Idisp, args...)
}

//ReadRow("A", 1, "F", 9  or "A", 1  or  1, 9  or  1  or  nothing, procfunc)
func (sheet Sheet) ReadRow(args... interface{}) {
    columnBegin, columnEnd, rowBegin, rowEnd, once := "", "", 0, 0, 20
    var proc func([]interface{}) int

    for _, arg := range args {
        switch arg.(type) {
            case int16:
                if n := int(arg.(int16)); n > 0 {
                    once = n
                }
            case func([]interface{}) int:
                proc = arg.(func([]interface{}) int)
            case string:
                if columnBegin == "" {
                    columnBegin = arg.(string)
                } else if columnEnd == "" {
                    columnEnd = arg.(string)
                }
            case int:
                if rowBegin <= 0 {
                    rowBegin = arg.(int)
                } else if rowEnd <= 0 {
                    rowEnd = arg.(int)
                }
        }
    }
    if proc == nil {
        panic("ReadRow proc is nil, want func([]interface{}) int")
    }

    if columnBegin == "" {
        columnBegin = "A"
    }
    if columnEnd == "" {
        if ucc, cb := int(sheet.MustGet("UsedRange", "Columns", "Count").(int32)), ColumnAtoi(columnBegin); ucc > cb {
            columnEnd = ColumnItoa(ucc)
        } else {
            columnEnd = columnBegin
        }
    }
    if rowBegin <= 0 {
        rowBegin = 1
    }
    if rowEnd <= 0 {
        if urc := int(sheet.MustGet("UsedRange", "Rows", "Count").(int32)); urc > rowBegin {
            rowEnd = urc
        } else {
            rowEnd = rowBegin
        }
    }

    for rbi, rei := rowBegin, rowBegin - 1; rei < rowEnd; rbi = rei + 1 {
        if rei += once; rei > rowEnd {
            rei = rowEnd
        }
        val := sheet.MustGetRange(fmt.Sprintf("%v%v:%v%v", columnBegin, rbi, columnEnd, rei))
        if v, ok := val.([][]interface{}); ok {
            for _, row := range v {
                if rc := proc(row); rc == -1 {
                    goto END
                }
            }
        } else if rc := proc([]interface{} {val}); rc == -1 {
            goto END
        }
    }
END:
}

//put range Property.
func (rg Range) Put(args... interface{}) (error) {
    return PutProperty(rg.Idisp, args...)
}

//get range Property as interface.
func (rg Range) Get(args... string) (interface{}, error) {
    return GetProperty(rg.Idisp, args...)
}

//
func (rg Range) MustGet(args... string) (interface{}) {
    return MustGetProperty(rg.Idisp, args...)
}

//
func (rg Range) Release() {
    rg.Idisp.Release()
}

//get Property as interface.
func (cell Cell) Get(args... string) (interface{}, error) {
    return GetProperty(cell.Idisp, args...)
}

//Must get Property as interface.
func (cell Cell) MustGet(args... string) (interface{}) {
    return MustGetProperty(cell.Idisp, args...)
}

//get Property as string.
func (cell Cell) Gets(args... string) (string, error) {
    v, err := cell.Get(args...)
    return String(v), err
}

//Must get Property as string.
func (cell Cell) MustGets(args... string) (ret string) {
    return String(cell.MustGet(args...))
}

//put cell Property.
func (cell Cell) Put(args... interface{}) (error) {
    return PutProperty(cell.Idisp, args...)
}

//
func (cell Cell) Release() {
    cell.Idisp.Release()
}

//get Property as interface.
func GetProperty(idisp *ole.IDispatch, args... string) (ret interface{}, err error) {
    defer Except("GetProperty", &err)
    argnum := len(args)
    if argnum==0 {
        ret = VARIANT{oleutil.MustGetProperty(idisp, "Value")}.Value()
    } else {
        maxi := argnum - 1
        for i:=0; i<maxi && err==nil; i++ {
            idisp = oleutil.MustGetProperty(idisp, args[i]).ToIDispatch()
            defer DoFuncs(idisp.Release)
        }
        //get multi-Property
        argv := args[maxi]
        if strings.IndexAny(argv, ",") != -1 {
            sl := []string{}
            for _, key := range strings.Split(argv, ",") {
                sl = append(sl, key+":"+String(VARIANT{oleutil.MustGetProperty(idisp, key)}.Value()))
            }
            ret = strings.Join(sl, ", ")
        } else {
            ret = VARIANT{oleutil.MustGetProperty(idisp, argv)}.Value()
        }
    }
    return
}

//get Property as interface.
func MustGetProperty(idisp *ole.IDispatch, args... string) (interface{}) {
    ret, err := GetProperty(idisp, args...)
    if err != nil {
        panic(err)
    }
    return ret
}

//put Property.
func PutProperty(idisp *ole.IDispatch, args... interface{}) (err error) {
    defer Except("PutProperty", &err)
    argnum := len(args)
    if argnum==1 {
        oleutil.MustPutProperty(idisp, "Value", args[0])
    } else if argnum>1 {
        maxi := argnum-2
        for i:=0; i<maxi && err==nil; i++ {
            idisp = oleutil.MustGetProperty(idisp, args[i].(string)).ToIDispatch()
            defer DoFuncs(idisp.Release)
        }
        //put multi-Property
        if argv, ok := args[argnum-1].(map[string]interface{}); ok {
            idisp = oleutil.MustGetProperty(idisp, args[maxi].(string)).ToIDispatch()
            defer DoFuncs(idisp.Release)
            for key, val := range argv {
                oleutil.MustPutProperty(idisp, key, val)
            }
        } else {
            oleutil.MustPutProperty(idisp, args[maxi].(string), args[argnum-1])
        }
    } else {
        err = errors.New("args is empty")
    }
    return
}

//get val of MS VARIANT
func (va VARIANT) Value() (val interface{}) {
    switch va.VT {
        case 0:               //VT_EMPTY
            val = ""
        case 1:               //VT_NULL
            val = ""
        case 2:
            val = *((*int16)(unsafe.Pointer(&va.Val)))
        case 3:
            val = *((*int32)(unsafe.Pointer(&va.Val)))
        case 4:
            val = *((*float32)(unsafe.Pointer(&va.Val)))
        case 5:
            val = *((*float64)(unsafe.Pointer(&va.Val)))
        case ole.VT_CY:
            _val6 := *((*int64)(unsafe.Pointer(&va.Val)))
            val = float64(_val6)/10000
        case 7:               //VT_DATE. Unix(second):1970-1-1. excel(day):1900-1-1. 19 for leap year. 0.5 for round.
            _val7 := *((*float64)(unsafe.Pointer(&va.Val)))
            val = time.Unix(int64((_val7-70*365-19)*24*3600+0.5), 0).Format("2006-01-02 15:04:05")
        case 8:               //string
            _val8 := *((**uint16)(unsafe.Pointer(&va.Val)))
            val = ole.UTF16PtrToString(_val8)
        case 9:               //*IDispatch
            val = va
        case 10:              //VT_ERROR
            val = "#VT_ERROR"
        case 11:
            val = *((*bool)(unsafe.Pointer(&va.Val)))
        //case 12:              //VT_VARIANT
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
        case ole.VT_ARRAY, 0x200c:  //8204,range get, 0x2000(VT_ARRAY) + 0xC(VT_VARIANT)
            val = ToValueArray(va.ToArray())
            //val = va.ToArray().ToValueArray()
        default:
            val = va
    }
    return
}

//
func String(val interface{}) (ret string) {
    switch val.(type) {
        case int:
            ret = strconv.FormatInt(int64(val.(int)), 10)
        case int8:
            ret = strconv.FormatInt(int64(val.(int8)), 10)
        case int16:
            ret = strconv.FormatInt(int64(val.(int16)), 10)
        case int32:
            ret = strconv.FormatInt(int64(val.(int32)), 10)
        case int64:
            ret = strconv.FormatInt(val.(int64), 10)
        case float32:
            ret = strconv.FormatFloat(float64(val.(float32)), 'f', -1, 64)
        case float64:
            ret = strconv.FormatFloat(val.(float64), 'f', -1, 64)
        case uint8:
            ret = strconv.FormatUint(uint64(val.(uint8)), 10)
        case uint16:
            ret = strconv.FormatUint(uint64(val.(uint16)), 10)
        case uint32:
            ret = strconv.FormatUint(uint64(val.(uint32)), 10)
        case uint64:
            ret = strconv.FormatUint(val.(uint64), 10)
        case *uint16:                     //string
            ret = ole.UTF16PtrToString(val.(*uint16))
        case bool:
            if val.(bool) {
                ret = "true"
            } else {
                ret = "false"
            }
        case string:
            ret = val.(string)
        default:
            ret = fmt.Sprintf("%+v", val)
    }
    return
}

//
func ColumnItoa(num int) (col string) {
    for num -= 1; num >= 0; num = num / 26 - 1 {
        d := num % 26
        col = string('A' + d) + col
    }
    return
}

//
func ColumnAtoi(s string) int {
    num, b, s := 0, 1, strings.ToUpper(s)
    for i := len(s) - 1; i >= 0; i -- {
        num += b * (int(s[i]) - 'A' + 1)
        b = b * 26
    }
    return num
}

//
func Except(info string, err *error, funcs... interface{}) {
    r := recover()
    if err != nil {
        if *err != nil {
            fmt.Fprintf(os.Stderr, "*%v: %+v\n", info, *err)
        }
        if r != nil {
            e := fmt.Sprintf("@%v: %+v", info, r)
            fmt.Fprintf(os.Stderr, "%v\n", e)
            *err = errors.New(e)
            debug.PrintStack()
        } else if *err != nil {
            *err = errors.New("%"+info+"%"+(*err).Error())
        }
    }
    if funcs != nil {
        DoFuncs(funcs...)
    }
}

//
func DoFuncs(funcs... interface{}) {
    if len(funcs) != 0 {
        fx, args := reflect.Value{}, []reflect.Value{}
        for _, one := range funcs {
            cur := reflect.ValueOf(one)
            if cur.Kind().String() == "func" {
                if fx.IsValid() {
                    RftCall(fx, args...)
                }
                fx, args = cur, []reflect.Value{}
            } else {
                args = append(args, cur)
            }
        }
        if fx.IsValid() {
            RftCall(fx, args...)
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

//from github.com/go-ole/go-ole/utility.go:convertHresultToError
// convertHresultToError converts syscall to error, if call is unsuccessful.
func convertHresultToError(hr uintptr, r2 uintptr, ignore error) (err error) {
    if hr != 0 {
        err = ole.NewError(hr)
    }
    return
}

//from github.com/go-ole/go-ole/safearray_windows.go:safeArrayGetVartype
// AKA: SafeArrayGetVartype in Windows API.
func safeArrayGetVartype(safearray *ole.SafeArray) (varType uint16, err error) {
    err = convertHresultToError(
        procSafeArrayGetVartype.Call(
            uintptr(unsafe.Pointer(safearray)),
            uintptr(unsafe.Pointer(&varType))))
    return
}

//from github.com/go-ole/go-ole/safearray_windows.go:safeArrayGetElement
// safeArrayGetElement retrieves element at given index.
func safeArrayGetElement(safearray *ole.SafeArray, index [2]int32, pv unsafe.Pointer) error {
    indexPtr := unsafe.Pointer(&index[0])
    return convertHresultToError(
        procSafeArrayGetElement.Call(
            uintptr(unsafe.Pointer(safearray)),
            uintptr(indexPtr),
            uintptr(pv)))
}

//from github.com/go-ole/go-ole/safearrayconversion.go:ToValueArray
func ToValueArray(sac *ole.SafeArrayConversion) (values [][]interface{}) {
    totalElements1, _ := sac.TotalElements(1)
    totalElements2, _ := sac.TotalElements(2)
    te1, te2 := int(totalElements1), int(totalElements2)

    values = make([][]interface{}, te1)
    for i := 0; i < te1; i ++ {
        row := make([]interface{}, te2)
        for j := 0; j < te2; j ++ {
            var v ole.VARIANT
            safeArrayGetElement(sac.Array, [2]int32 {int32(i)+1, int32(j)+1}, unsafe.Pointer(&v))
            row[j] = (VARIANT{&v}).Value()
        }
        values[i] = row
    }

    return
}

