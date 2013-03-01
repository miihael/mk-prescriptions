package main

import (
    "fmt"
    "time"
)

func main() {
    t := time.Date(1900, 1, 1, 0, 0, 0, 0, time.UTC)
    t = t.AddDate(0, 0, 41330 - 2 )
    fmt.Printf("result: %04d-%02d-%02d\n", t.Year(), t.Month(), t.Day())
}

