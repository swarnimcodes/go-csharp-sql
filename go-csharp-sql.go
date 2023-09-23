package main

import (
	"bufio"
	"fmt"
	"io/fs"
	"log"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/xuri/excelize/v2"
)

func ret_sp_name(file_line string) (string, error) {
	pattern := `"[^"]+"`
	re, err := regexp.Compile(pattern)
	if err != nil {
		fmt.Printf("Error compiling regexp: %s", err)
		return "", err
	}

	match := re.FindString(file_line)

	if match != "" {
		match = match[1 : len(match)-1]
		return match, err
	} else {
		return "", err
	}
}

func ret_rec_filelist(path string) ([]string, error) {
	var filelist []string
	err := filepath.WalkDir(path, func(path string, d fs.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if !d.IsDir() {
			filelist = append(filelist, path)
		}

		return nil
	})

	if err != nil {
		return nil, err
	}

	return filelist, nil
}

func write_excel_sp(filelist []string, sp_methods []string, table_methods []string) error {
	xl := excelize.NewFile()

	// Headers
	xl.SetCellValue("Sheet1", "A1", "File Path")
	xl.SetCellValue("Sheet1", "B1", "SP Line No.")
	xl.SetCellValue("Sheet1", "C1", "SP Name")

	//not implemented
	xl.SetCellValue("Sheet1", "D1", "Table Query Line No.")
	xl.SetCellValue("Sheet1", "E1", "Table")

	row := 2

	for _, file := range filelist {
		if path.Ext(file) == ".cs" {
			file_content, err := os.Open(file)

			if err != nil {
				log.Fatalf("Error reading file: %s", err)
				continue
			}
			defer file_content.Close()

			scanner := bufio.NewScanner(file_content)

			ln := 1

			for scanner.Scan() {
				if !strings.HasPrefix(scanner.Text(), "//") {
					for _, sp_method := range sp_methods {
						if strings.Contains(scanner.Text(), sp_method) {
							sp_name, err := ret_sp_name(scanner.Text())
							if err != nil {
								log.Fatalf("Error returning SP Name: %s", err)
								continue
							}

							xl.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), file)
							xl.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), ln)
							xl.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), sp_name)
							row++
						}
					}
				}
				ln = ln + 1
			}
		}
	}

	if err := xl.SaveAs("cs_output.xlsx"); err != nil {
		log.Fatalf("Error saving Excel File: %s", err)
		return err
	}
	return nil
}

func main() {
	sp_methods := []string{"ExecuteNonQuerySP", "ExecuteDataSetSP"}
	table_methods := []string{"FillDropDownOnly"}
	fmt.Println("Hello, World!")

	cs_dir := "/home/swarnim/Downloads/backup/py-csharp-sql/cs"

	filelist, err := ret_rec_filelist(cs_dir)

	if err != nil {
		log.Fatalf("Error listing files: %s", err)
	}

	if err := write_excel_sp(filelist, sp_methods, table_methods); err != nil {
		log.Fatalf("Error writing to Excel: %s", err)
	}
}
