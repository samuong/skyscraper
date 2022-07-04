package main

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {
	progname := filepath.Base(os.Args[0])

	if len(os.Args) < 3 {
		fmt.Fprintf(os.Stderr, "usage: %s sheet.xlsx cellref[=val]...\n", progname)
		os.Exit(1)
	}

	f, err := excelize.OpenFile(os.Args[1])
	if err != nil {
		fmt.Fprintf(os.Stderr, "error: %s\n", err)
		os.Exit(1)
	}
	defer f.Close()

	for _, arg := range os.Args[2:] {
		lvalue, rvalue, found := strings.Cut(arg, "=")
		if found {
			// Contains '=' token; write rvalue to the referenced cell.
			if err := f.SetCellValue("Sheet1", lvalue, rvalue); err != nil {
				fmt.Fprintf(os.Stderr, "%s: warning: %s\n", progname, err)
				continue
			}
		} else {
			// No '=' token; read value from the cell reference.
			val, err := f.CalcCellValue("Sheet1", lvalue)
			if err != nil {
				fmt.Fprintf(os.Stderr, "%s: warning: %s\n", progname, err)
				continue
			}
			fmt.Println(val)
		}
	}
}
