package main

import (
    "encoding/xml"
    "fmt"
    "os"
    "io"
    "io/ioutil"
    "path"
    "archive/zip"
    "log"
    "time"
    "strings"
    "strconv"
    "bytes"
)

type Prescription struct {
    Num int
    Date string
    Day int
    Mon string
    Year int
    Fio string
    Goal string
    Addr string
    Vch string
    Pos string
}

func ParseSharedStrings(content []byte) []string {
    type Result struct {
        XMLName xml.Name `xml:"sst"`
        Count int `xml:"count,attr"`
        Strings []string `xml:"si>t"`
    }
    v := Result{}
    err := xml.Unmarshal(content, &v);
    if err != nil {
        fmt.Printf("error: %v", err);
        return nil;
    }
    return v.Strings;
}

func RusMonth(mon int) string {
    switch mon {
    case 1: return "января"
    case 2: return "февраля"
    case 3: return "марта"
    case 4: return "апреля"
    case 5: return "мая"
    case 6: return "июня"
    case 7: return "июля"
    case 8: return "августа"
    case 9: return "сентября"
    case 10: return "октября"
    case 11: return "ноября"
    case 12: return "декабря"
    }
    return ""
}

func MakeDate(excel_date int) (date string, day int, mon string, year int) {
    t := time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)
    t = t.AddDate(0, 0, excel_date - 2 )
    //fmt.Printf("result: %s\n", t.UTC())
    date = fmt.Sprintf("%02d%02d%02d\n", t.Day(), t.Month(), t.Year()%100)
    day = t.Day()
    mon = RusMonth(int(t.Month()))
    year = t.Year()
    return date, day, mon, year
}

func ParseSheet(content []byte, strings []string) []Prescription {
    type Col struct {
        Val int `xml:"v"`
        Type string `xml:"t,attr"`
    }

    type Row struct {
        Num int `xml:"r,attr"`
        Cols []Col `xml:"c"`
    }

    type Result struct {
        XMLName xml.Name `xml:"worksheet"`
        Rows []Row `xml:"sheetData>row"`
    }
    v := Result{};
    err := xml.Unmarshal(content, &v);
    if err != nil {
        fmt.Printf("error: %v\n", err);
        return nil;
    }
    prps := make([]Prescription, 10)
    for _, row := range v.Rows {
        fmt.Printf("%d:", row.Num)
        p := Prescription{}
        for j, col := range row.Cols {
            if j==0 && col.Val==0 { continue; }
            if col.Type == "s" {
               v := strings[col.Val]
               if row.Num > 1 {
                    switch j {
                    case 2: p.Fio = v
                    case 3: p.Goal = v
                    case 4: p.Addr = v
                    case 5: p.Vch = v
                    case 6: p.Pos = v
                    }
                }
                fmt.Printf("'%v' ", strings[col.Val]);
            } else {
                if row.Num > 1 {
                    switch j {
                    case 0: p.Num = col.Val
                    case 1: p.Date, p.Day, p.Mon,p.Year = MakeDate(col.Val)
                    }
                }
                fmt.Printf("%d ", col.Val);
            }
        }
        if p.Num > 0 { prps = append(prps, p); }
        fmt.Printf("\n")
    }
    return prps;
}

func Unpack(filename string, prefix string) {
    // Open a zip archive for reading.
    os.RemoveAll(prefix)
    r, err := zip.OpenReader(filename)
    if err != nil {
        log.Fatal(err)
    }
    defer r.Close()

    flist, lerr := os.Create(prefix+".list")
    if lerr != nil {
        log.Fatal(lerr)
    }

    // Iterate through the files in the archive,
    // printing some of their contents.
    for _, f := range r.File {
        if f.CompressedSize <= 0 { continue }
        fmt.Fprintf(flist, "%s\n", f.Name)
        err := os.MkdirAll(path.Dir(prefix + "/" + f.Name), 0755)
        if err != nil {
            log.Fatal(err)
        }
        rc, err := f.Open()
        if err != nil {
            log.Fatal(err)
        }
        dstf, err := os.Create(prefix + "/" + f.Name)
        if err!= nil {
            log.Fatal(err)
        }
        defer dstf.Close()
        _, err = io.Copy(dstf, rc)
        if err != nil {
            log.Fatal(err)
        }
        rc.Close()
        //fmt.Println()
    }
    flist.Close()
}

func Pack(listfile string, outfile string) {
    listbuf, err := ioutil.ReadFile(listfile)
    if err!=nil { log.Fatal(err) }
    list := bytes.Split(listbuf, []byte("\n"))

    buf, err := os.Create(outfile)
    if err!=nil || len(list)==0 {
        log.Fatal(fmt.Sprintf("can't write %s", outfile))
    }
    // Create a new zip archive.

    w := zip.NewWriter(buf)
    for _, file := range list {
        if len(file)<=1 { continue }
        f, err := w.Create(string(file))
        if err != nil {
            log.Fatal(err)
        }
        //fmt.Println(string(file))
        srcf, err := os.Open("blank/" + string(file))
        if err != nil { log.Fatal(err) }
        defer srcf.Close()
        _, err = io.Copy(f, srcf)
        if err != nil { log.Fatal(err) }
    }

    // Make sure to check the error on Close.
    err = w.Close()
    if err != nil {
        log.Fatal(err)
    }
}

func OutputPrp(template []byte, ps ... Prescription) {
    os.MkdirAll("output", 0755)
    p := ps[0]
    fnum := strconv.Itoa(p.Num)
    args := []string{
             "#{NUM1}", strconv.Itoa(p.Num), "#{DATE1}", p.Date,
             "#{DAY}", strconv.Itoa(p.Day), "#{MON}", p.Mon, "#{YEAR}", strconv.Itoa(p.Year),
             "#{FIO1}", p.Fio, "#{POS1}", p.Pos, "#{ADDR1}", p.Addr,
             "#{VCH1}", p.Vch, "#{GOAL1}", p.Goal }
    if len(ps)>1 {
        p = ps[1]
        fnum += "_" + strconv.Itoa(p.Num)
        args = append(args,
             "#{NUM2}", strconv.Itoa(p.Num), "#{DATE2}", p.Date,
             "#{FIO2}", p.Fio, "#{POS2}", p.Pos, "#{ADDR2}", p.Addr,
             "#{VCH2}", p.Vch, "#{GOAL2}", p.Goal )
    }
    outf, err := os.Create("blank/word/document.xml")
    if err != nil { log.Fatal(err); }
    r := strings.NewReplacer(args...)
    r.WriteString(outf, string(template))
    outf.Close();
    docx := fmt.Sprintf("output/p_%s.docx", fnum)
    fmt.Printf("saving %s\n", docx)
    Pack("blank.list", docx )
}

func main() {
    Unpack("form.xlsx", "form");
    sst_path := "form/xl/sharedStrings.xml";
    content, err := ioutil.ReadFile(sst_path);
    if err != nil {
        fmt.Printf("Error opening file %s: %s\n", sst_path, err);
        return;
    }
    //fmt.Printf("%s\n\n", content);
    strings := ParseSharedStrings(content);

    sh_path := "form/xl/worksheets/sheet1.xml"
    content, err = ioutil.ReadFile(sh_path);
    if err != nil { log.Fatal(err) }
    prps := ParseSheet(content, strings);
    Unpack("blank_template.docx", "blank")
    fmt.Println("making prescriptions")
    template, err := ioutil.ReadFile("blank/word/document.xml")
    if err != nil {
        log.Fatal(err);
    }
    for i, p := range prps[:len(prps)-1] {
        if (p.Num > 0) && (i%2) == 0 { OutputPrp(template, p, prps[i+1]) }
    }
    if len(prps)%2 == 1 { OutputPrp(template, prps[len(prps)-1]) }
    fmt.Println("finished")
}
