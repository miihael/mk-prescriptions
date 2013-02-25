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
)

type Worker struct {
    DocNum int
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

func ParseSheet(content []byte, strings []string) []Worker {
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
    for _, row := range v.Rows {
        fmt.Printf("%d:", row.Num)
        for _, col := range row.Cols {
            if col.Type == "s" {
                fmt.Printf("'%v' ", strings[col.Val]);
            } else {
                fmt.Printf("%d ", col.Val);
            }
        }
        fmt.Printf("\n")
    }
    return nil;
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
    defer flist.Close()

    // Iterate through the files in the archive,
    // printing some of their contents.
    for _, f := range r.File {
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
}


func main() {
    Unpack("form.xlsx", "form");
    sst_path := "form/xl/sharedStrings.xml";
    content, err := ioutil.ReadFile(sst_path);
    if err != nil {
        fmt.Printf("error opening file %s: %s\n", sst_path, err);
        return;
    }
    //fmt.Printf("%s\n\n", content);
    strings := ParseSharedStrings(content);

    sh_path := "form/xl/worksheets/sheet1.xml"
    content, err = ioutil.ReadFile(sh_path);
    if err != nil {
        fmt.Printf("error opening file %s: %s\n", sh_path, err);
        return;
    }
    workers := ParseSheet(content, strings);
    fmt.Printf("%+v", workers);
    Unpack("blank.docx", "blank")
}
