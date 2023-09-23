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

func returnTableNames(tablequery string) []string {

	pattern := `\btbl\w+`
	re := regexp.MustCompile(pattern)
	matches := re.FindAllString(tablequery, -1)

	var filteredMatches []string
	for _, match := range matches {
		match = strings.TrimSpace(match)
		if match != "" {
			filteredMatches = append(filteredMatches, match)
		}
	}
	return filteredMatches
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
	xl.SetCellValue("Sheet1", "B1", "SP Count")
	xl.SetCellValue("Sheet1", "C1", "SP Line No.")
	xl.SetCellValue("Sheet1", "D1", "SP List")
	xl.SetCellValue("Sheet1", "E1", "Table Count")
	xl.SetCellValue("Sheet1", "F1", "Table List")
	xl.SetCellValue("Sheet1", "G1", "Query Line No.")

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
			var spLn []string
			var tblLn []string

			var spList, tableList []string

			for scanner.Scan() {
				line := scanner.Text()
				line = strings.TrimSpace(line)
				if strings.HasPrefix(line, "//") {
					ln++
					continue
				}

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
					spList = append(spList, sp_name)
					spLn = append(spLn, fmt.Sprintf("%d", ln))
				}

				if sp_or_table == "Table" {
					tableNames := returnTableNames(scanner.Text())
					tableList = append(tableList, tableNames...)
					tblLn = append(tblLn, fmt.Sprintf("%d", ln))
				}
				ln++
			}
			// Set cell values for this file
			xl.SetCellValue("Sheet1", fmt.Sprintf("A%d", row), file)
			xl.SetCellValue("Sheet1", fmt.Sprintf("B%d", row), len(spList))
			xl.SetCellValue("Sheet1", fmt.Sprintf("C%d", row), strings.Join(spLn, "\n")) //TODO: SP Line Number
			xl.SetCellValue("Sheet1", fmt.Sprintf("D%d", row), strings.Join(spList, "\n"))
			xl.SetCellValue("Sheet1", fmt.Sprintf("E%d", row), len(tableList))
			xl.SetCellValue("Sheet1", fmt.Sprintf("F%d", row), strings.Join(tableList, "\n"))
			xl.SetCellValue("Sheet1", fmt.Sprintf("G%d", row), strings.Join(tblLn, "\n")) //TODO: Table Line Numbers

			row++
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
