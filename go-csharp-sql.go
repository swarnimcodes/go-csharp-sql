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

func spMethodList() []string {
	return []string{
		"ExecuteNonQuery",
		"ExecuteDataSet",
		"ExecuteNonQueryAsync",
		"ExecuteReader",
		"ExecuteReaderAsync",
		"ExecuteScalar",
		"ExecuteScalarAsync",
	}
}

func tableMethodsList() []string {
	return []string{
		"FillDropDownOnly",
	}
}

func containsString(line string, list []string) bool {
	for _, item := range list {
		if strings.Contains(line, item) {
			return true
		}
	}
	return false
}

func isSPOrTable(line string) (string, error) {
	if containsString(line, spMethodList()) {
		return "SP", nil
	} else if containsString(line, tableMethodsList()) {
		return "Table", nil
	}
	return "", fmt.Errorf("No match found!")
}

func returnSPName(file_line string) (string, error) {
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

func returnTableNames(tablequery string) string {

	pattern := `\btbl\w+`
	re := regexp.MustCompile(pattern)
	matches := re.FindAllString(tablequery, -1)

	tableNames := strings.Join(matches, "\n")

	return tableNames
}

func returnRecursiveFilelist(path string) ([]string, error) {
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

func writeToExcel(filelist []string) error {
	xl := excelize.NewFile()

	// Headers
	xl.SetCellValue("Sheet1", "A1", "File Path")
	xl.SetCellValue("Sheet1", "B1", "SP Line No.")
	xl.SetCellValue("Sheet1", "C1", "SP Name")
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
				line := scanner.Text()
				line = strings.TrimSpace(line)
				if strings.HasPrefix(line, "//") {
					ln++
					continue
				}

				// TODO: delegate checking to outside func

				sp_or_table, err := isSPOrTable(scanner.Text())
				if err != nil {
					// log.Fatalf("%s", err)
					ln++
					continue
				}

				if sp_or_table == "SP" {
					sp_name, err := returnSPName(scanner.Text())
					if err != nil {
						log.Fatalf("Error returning SP Name: %s", err)
						continue
					}
					xl.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), file)
					xl.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), ln)
					xl.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), sp_name)
					row++
				}

				if sp_or_table == "Table" {
					tableNames := returnTableNames(scanner.Text())
					xl.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), file)
					xl.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), "")
					xl.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), "")
					xl.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), ln)
					xl.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), tableNames)
					row++
				}

				ln++
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
	fmt.Println("Hello, World!")

	cs_dir := "/home/swarnim/Downloads/backup/py-csharp-sql/cs"

	filelist, err := returnRecursiveFilelist(cs_dir)

	if err != nil {
		log.Fatalf("Error listing files: %s", err)
	}

	if err := writeToExcel(filelist); err != nil {
		log.Fatalf("Error writing to Excel: %s", err)
	}
}
