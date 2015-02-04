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
    date = fmt.Sprintf("%02d.%02d.%02d\n", t.Day(), t.Month(), t.Year()%100)
    day = t.Day()
    mon = RusMonth(int(t.Month()))
    year = t.Year()
    return date, day, mon, year
}

type Cell struct {
    S string
    N int
}

func ParseSheet(content []byte, strings []string) [][]Cell {
    type Col struct {
        Val string `xml:"v"`
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
    data := make([][]Cell, len(v.Rows))
    for i, row := range v.Rows {
        data[i] = make([]Cell, len(row.Cols))
        fmt.Printf("row %d: ", row.Num)
        for j, col := range row.Cols {
            if col.Type == "s" { //shared string
               idx, err := strconv.Atoi(col.Val);
               if err == nil {
                   fmt.Printf("%d ", idx);
                   data[i][j].S = strings[idx]
                   fmt.Printf("'%v' ", strings[idx]);
               } else {
                    data[i][j].S = col.Val;
               }
            } else {
              //case 1: p.Date, p.Day, p.Mon,p.Year = MakeDate(col.Val)
               data[i][j].N, err = strconv.Atoi(col.Val)
               if err != nil {
                  data[i][j].S = col.Val
               }
               fmt.Printf("%s ", col.Val);
            }
        }
        fmt.Printf("\n")
    }
    return data;
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

func MakeArgs(p []Cell, fchar string) []string {
    alen := len(p)*2
    args := make([]string, alen)
    for i, cell := range p {
        args[i*2] = fmt.Sprintf("%s%02d%s", fchar, i+1, fchar)
        if cell.N > 0 {
            if cell.N > 30000 &&  cell.N < 44000 {
                dt, day, mon, y := MakeDate(cell.N)
                args[i*2+1] = dt
                args = append(args, []string{fmt.Sprintf("%s%02d@%s", fchar, i+1, fchar), strconv.Itoa(day)}...)
                args = append(args, []string{fmt.Sprintf("%s%02d$%s", fchar, i+1, fchar), mon}...)
                args = append(args, []string{fmt.Sprintf("%s%02d*%s", fchar, i+1, fchar), strconv.Itoa(y)}...)
            } else {
                args[i*2+1] = strconv.Itoa(cell.N)
            }
        } else {
            args[i*2+1] = cell.S
        }
    }
    fmt.Printf("%v\n", args);
    return args
}

func OutputReport(template []byte, p []Cell, s []Cell) {
    os.MkdirAll("output", 0755)
    fnum := fmt.Sprintf("%d-%d", p[0].N, s[0].N)

    outf, err := os.Create("blank/word/document.xml")
    if err != nil { log.Fatal(err); }

    args := MakeArgs(p, "##")
    if s!=nil {
        args = append(args, MakeArgs(s, "^^")...)
    }
    r := strings.NewReplacer(args...)
    r.WriteString(outf, string(template))
    outf.Close();
    docx := fmt.Sprintf("output/s-%s.docx", fnum)
    fmt.Printf("saving %s\n", docx)
    Pack("blank.list", docx)
}

func main() {
    if len(os.Args)<2 {
        fmt.Printf("usage: %s input_form.xlsx\n", os.Args[0])
        return
    }
    Unpack(os.Args[1], "iform")
    sst_path := "iform/xl/sharedStrings.xml";
    content, err := ioutil.ReadFile(sst_path);
    if err != nil {
        fmt.Printf("Error opening file %s: %s\n", sst_path, err);
        return;
    }
    //fmt.Printf("%s\n\n", content);
    strings := ParseSharedStrings(content);
    fmt.Printf("%v", strings);

    sh_path := "iform/xl/worksheets/sheet1.xml"
    content, err = ioutil.ReadFile(sh_path);
    if err != nil { log.Fatal(err) }
    data := ParseSheet(content, strings);
    fmt.Printf("%v\n", data)

    Unpack("t.docx", "blank")
    fmt.Println("making reports")
    template, err := ioutil.ReadFile("blank/word/document.xml")
    if err != nil {
        log.Fatal(err);
    }
    for i:=1; i<len(data)-1; i+=2 {
        fmt.Printf("i=%d\n", i)
        if data[i][0].N > 0 {
            OutputReport(template, data[i], data[i+1])
        }
    }
    fmt.Println("finished")
}
