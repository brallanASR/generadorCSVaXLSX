package main

import (
	"encoding/csv"
	"fmt"
	"io"
	"os"
	"strings"

	excelize "github.com/xuri/excelize/v2"
)

func main() {
	files, _ := os.ReadDir(".")
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
	r.Comma = detectSeparator(f)
	r.LazyQuotes = true

	xlsx := excelize.NewFile()
	sheet := "Sheet1"

	rowNum := 1
	for {
		record, err := r.Read()
		if err == csv.ErrFieldCount {
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

	outputFilename := strings.TrimSuffix(csvFile, ".csv") + ".xlsx"
	err = xlsx.SaveAs(outputFilename)
	if err != nil {
		fmt.Println("Error al guardar XLSX:", err)
		return
	}
	fmt.Printf("Converted: %s -> %s\n", csvFile, outputFilename)
}

func detectSeparator(file *os.File) rune {
	reader := csv.NewReader(file)

	reader.Comma = ';'
	if _, err := reader.Read(); err == nil {
		file.Seek(0, io.SeekStart)
		return ';'
	}

	reader.Comma = ','
	if _, err := reader.Read(); err == nil {
		file.Seek(0, io.SeekStart)
		return ','
	}

	file.Seek(0, io.SeekStart)
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
