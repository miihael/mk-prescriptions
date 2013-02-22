package main

import (
    "archive/zip"
    "fmt"
    "log"
    "io"
    "os"
)

func read_zip(filename string) {
    // Open a zip archive for reading.
    r, err := zip.OpenReader(filename)
    if err != nil {
        log.Fatal(err)
    }
    defer r.Close()

    // Iterate through the files in the archive,
    // printing some of their contents.
    for _, f := range r.File {
        fmt.Printf("Contents of %s:\n", f.Name)
        rc, err := f.Open()
        if err != nil {
            log.Fatal(err)
        }
        _, err = io.CopyN(os.Stdout, rc, 20)
        if err != nil {
            log.Fatal(err)
        }
        rc.Close()
        fmt.Println()
    }
}

func write_zip(filename string) {
    buf, err := os.Create(filename)
    if err!=nil {
        log.Fatal(err)
    }
    // Create a new zip archive.
    w := zip.NewWriter(buf)

    // Add some files to the archive.
    var files = []struct {
        Name, Body string
    }{
        {"1/readme.txt", "This archive contains some text files."},
        {"2/gopher.txt", "Gopher names:\nGeorge\nGeoffrey\nGonzo"},
        {"2/todo.txt", "Get animal handling licence.\nWrite more examples."},
    }
    for _, file := range files {
        f, err := w.Create(file.Name)
        if err != nil {
            log.Fatal(err)
        }
        _, err = f.Write([]byte(file.Body))
        if err != nil {
            log.Fatal(err)
        }
    }

    // Make sure to check the error on Close.
    err = w.Close()
    if err != nil {
        log.Fatal(err)
    }
}

func main() {
    write_zip("test.zip")
    read_zip("test.zip")
}

