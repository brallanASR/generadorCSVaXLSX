package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"io/ioutil"
	"os"
	"strings"
	"time"

	excelize "github.com/xuri/excelize/v2"
)

func main() {
	files, _ := ioutil.ReadDir(".")
	for _, file := range files {
		filename := file.Name()
		if strings.HasSuffix(filename, ".csv") {
			convertToXlsx(filename)
		}
	}
}

func convertToXlsx(csvFile string) {
	f, err := os.Open(csvFile)
	if err != nil {
		fmt.Println("Error al abrir el archivo:", err)
		return
	}
	defer f.Close()

	r := csv.NewReader(f)
	r.Comma = detectSeparator(csvFile)
	r.LazyQuotes = true

	xlsx := excelize.NewFile()
	sheet := "Sheet1"
	xlsx.NewSheet(sheet)

	rowNum := 1
	for {
		record, err := r.Read()
		if err == csv.ErrFieldCount {
			// Ignorar líneas problemáticas
			continue
		} else if err == io.EOF {
			break
		} else if err != nil {
			fmt.Println("Error al leer el CSV:", err)
			return
		}

		for j, val := range record {
			cell := columnToAlpha(j) + fmt.Sprint(rowNum)
			xlsx.SetCellValue(sheet, cell, val)
		}
		rowNum++
	}

	outputFilename := fmt.Sprintf("documento convertido %s %s.xlsx", strings.TrimSuffix(csvFile, ".csv"), time.Now().Format("2006-01-02"))
	err = xlsx.SaveAs(outputFilename)
	if err != nil {
		fmt.Println("Error al guardar XLSX:", err)
		return
	}
	fmt.Printf("Converted: %s -> %s\n", csvFile, outputFilename)
}

func detectSeparator(filename string) rune {
	file, _ := os.Open(filename)
	defer file.Close()

	reader := csv.NewReader(file)
	reader.Comma = ';'
	if _, err := reader.Read(); err == nil {
		return ';'
	}

	reader.Comma = ':'
	if _, err := reader.Read(); err == nil {
		return ':'
	}

	return '|'
}

// Convert a column number to its corresponding Excel column label
func columnToAlpha(col int) string {
	result := ""
	for col >= 0 {
		result = string('A'+(col%26)) + result
		col = col/26 - 1
	}
	return result
}
