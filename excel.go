package excel

import (
    "strconv"
    "strings"
    "path/filepath"
    "unsafe"
    "reflect"
    "errors"
    //"fmt"
    "github.com/mattn/go-ole"
    "github.com/mattn/go-ole/oleutil"
)

type Option struct {
    Visible                      bool
    DisplayAlerts                bool
    ScreenUpdating               bool
}

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

//
type VARIANT struct {
    *ole.VARIANT
}

//convert MS VARIANT to string
func (va VARIANT) ToString() (ret string) {
    switch va.VT {
        case 2:
            v2:=(*int16)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatInt(int64(*v2), 10)
        case 3:
            v3:=(*int32)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatInt(int64(*v3), 10)
        case 4:
            v4:=(*float32)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatFloat(float64(*v4), 'f', 2, 64)
        case 5:
            v5:=(*float64)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatFloat(*v5, 'f', 2, 64)
        case 8:       //string
            v8:=(**uint16)(unsafe.Pointer(&va.Val))
            ret = ole.UTF16PtrToString(*v8)
        case 11:
            v11:=(*bool)(unsafe.Pointer(&va.Val))
            if *v11 {
                ret = "TRUE"
            } else {
                ret = "FALSE"
            }
        case 16:
            v16:=(*int8)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatInt(int64(*v16), 10)
        case 17:
            v17:=(*uint8)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatUint(uint64(*v17), 10)
        case 18:
            v18:=(*uint16)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatUint(uint64(*v18), 10)
        case 19:
            v19:=(*uint32)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatUint(uint64(*v19), 10)
        case 20:
            v20:=(*int64)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatInt(int64(*v20), 10)
        case 21:
            v21:=(*uint64)(unsafe.Pointer(&va.Val))
            ret = strconv.FormatUint(uint64(*v21), 10)
    }
    return
}

//get fields of struct Option
func (option Option) Fields() (ret []string) {
    fields := reflect.Indirect(reflect.ValueOf(option)).Type()
    num := fields.NumField()
    for i:=0; i<num; i++ {
        ret = append(ret, fields.Field(i).Name)
    }
    return
}

//
func (mso *MSO) SetOption(option Option, args... int) {
    opts, curs := reflect.ValueOf(&mso.Option).Elem(), reflect.ValueOf(option)
    for _, key := range option.Fields() {
        opt, cur := opts.FieldByName(key), curs.FieldByName(key)
        optv, curv := opt.Bool(), cur.Bool()
        if (args != nil && args[0] > 0) || optv != curv {
            opt.SetBool(curv)
            oleutil.PutProperty(mso.IdExcel, key, curv)
        }
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

    option := Option{Visible: true, DisplayAlerts: true, ScreenUpdating: true}
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
    mso.WorkBookAdd()
    return
}

//
func Open(full string, options... Option) (mso *MSO, err error) {
    defer Except("Open", &err)
    mso = Init(options...)
    _, err = mso.WorkBookOpen(full)
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
    oleutil.MustCallMethod(mso.IdWorkBooks, "Close")
    oleutil.MustCallMethod(mso.IdExcel, "Quit")
    mso.IdWorkBooks.Release()
    mso.IdExcel.Release()
    mso.IuApp.Release()
    return
}

//
func (mso *MSO) Pick(workx string, id interface {}) (ret *ole.IDispatch, err error) {
    defer Except("Pick", &err)
    if id_int, ok := id.(int); ok {
        ret = oleutil.MustGetProperty(mso.IdExcel, workx, id_int).ToIDispatch()
    } else if id_str, ok := id.(string); ok {
        ret = oleutil.MustGetProperty(mso.IdExcel, workx, id_str).ToIDispatch()
    }
    return
}

//
func (mso *MSO) WorkBooksCount() (int) {
    return (int)(oleutil.MustGetProperty(mso.IdWorkBooks, "Count").Val)
}

//
func (mso *MSO) WorkBooks() (wbs WorkBooks) {
    num := mso.WorkBooksCount()
    for i:=1; i<=num; i++ {
        wbs = append(wbs, WorkBook{oleutil.MustGetProperty(mso.IdExcel, "WorkBooks", i).ToIDispatch(), mso})
    }
    return
}

//
func (mso *MSO) WorkBookAdd() (WorkBook, error) {
    _wb , err := oleutil.CallMethod(mso.IdWorkBooks, "Add")
    defer Except("WorkBookAdd", &err)
    return WorkBook{_wb.ToIDispatch(), mso}, err
}

//
func (mso *MSO) WorkBookOpen(full string) (WorkBook, error) {
    _wb, err := oleutil.CallMethod(mso.IdWorkBooks, "open", full)
    defer Except("WorkBookOpen", &err)
    return WorkBook{_wb.ToIDispatch(), mso}, err
}

//
func (mso *MSO) WorkBookActivate(id interface {}) (wb WorkBook, err error) {
    defer Except("WorkBookActivate", &err)
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
func (mso *MSO) SheetsCount() (int) {
    sheets := oleutil.MustGetProperty(mso.IdExcel, "Sheets").ToIDispatch()
    return (int)(oleutil.MustGetProperty(sheets, "Count").Val)
}

//
func (mso *MSO) Sheets() (sheets []Sheet) {
    num := mso.SheetsCount()
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
func (mso *MSO) SheetAdd(wb *ole.IDispatch, args... string) (Sheet, error) {
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
func (mso *MSO) SheetSelect(id interface {}) (sheet Sheet, err error) {
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

//get value as string/set
func (sheet Sheet) Cells(r int, c int, vals...interface{}) (ret string) {
    defer NoExcept()
    cell := oleutil.MustGetProperty(sheet.Idisp, "Range", Cell2r(r, c)).ToIDispatch()
    if vals == nil {
        //ret = oleutil.MustGetProperty(cell, "Value").ToString()
        ret = VARIANT{oleutil.MustGetProperty(cell, "Value")}.ToString()
    } else {
        oleutil.PutProperty(cell, "Value", vals[0])
    }
    return
}

//convert (5,27) to "AA5", xp+office2003 cant use cell(1,1)
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
func Except(info string, err *error, functions... func()) {
    r := recover()
    if r != nil {
        switch r.(type) {
            case string:
                info = "%"+info+":"+r.(string)
        }
        *err = errors.New(info)
    } else if err != nil && *err != nil {
        *err = errors.New("%"+info+":"+(*err).Error())
    }
    if functions != nil {
        NoExcept(functions...)
    }
}

//
func NoExcept(functions... func()) {
    recover()
    for _, funcx := range functions {
        func() {
            defer func() {recover()}()
            funcx()
        }()
    }
}





