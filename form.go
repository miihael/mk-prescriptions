package main

import (
    "encoding/xml"
    "fmt"
    //"os"
    "io/ioutil"
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
    type Row struct {
        Num int `xml:"r,attr"`
        Cols []int `xml:"c>v"`
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
            fmt.Printf("'%v' ", strings[col]);
        }
        fmt.Printf("\n")
    }
    return nil;
}

func main() {
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
}
