package main

import (
	"fmt"
	"log"
	"os"
	"sync"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	// 42.93s
	// generateExcelWithTempFiles()

	// 2.41s
	// generateExcelWithStreamWriter()

	// 11.06s
	// generateExcelConcurrent2()

	// 9.71s
	generateExcel()

	// readExcel()
}

func generateExcelWithTempFiles() {
	// เริ่มการวัดเวลา
	start := time.Now()

	var wg sync.WaitGroup
	numWorkers := 5
	rowsPerWorker := 200000

	// สร้างไฟล์ Excel ชั่วคราว
	tempFiles := make([]string, numWorkers)
	for i := 0; i < numWorkers; i++ {
		tempFile := fmt.Sprintf("temp_%d.xlsx", i)
		tempFiles[i] = tempFile

		wg.Add(1)
		go func(fileName string, startRow int) {
			defer wg.Done()

			f := excelize.NewFile()
			sheetName := "Sheet1"
			f.NewSheet(sheetName)

			// ใช้ StreamWriter สำหรับการเขียนข้อมูล
			streamWriter, err := f.NewStreamWriter(sheetName)
			if err != nil {
				log.Fatal(err)
			}

			for j := 0; j < rowsPerWorker; j++ {
				rowNum := startRow + j
				row := []interface{}{
					fmt.Sprintf("Row %d", rowNum),
					"Cell B",
				}
				cell, _ := excelize.CoordinatesToCellName(1, j+1)
				if err := streamWriter.SetRow(cell, row); err != nil {
					log.Fatal(err)
				}
			}

			if err := streamWriter.Flush(); err != nil {
				log.Fatal(err)
			}

			if err := f.SaveAs(fileName); err != nil {
				log.Fatal(err)
			}
		}(tempFiles[i], i*rowsPerWorker+1)
	}

	wg.Wait()

	// รวมไฟล์ Excel ชั่วคราว
	finalFile := "LargeFile.xlsx"
	fFinal := excelize.NewFile()
	sheetName := "Sheet1"
	fFinal.NewSheet(sheetName)

	for i, tempFile := range tempFiles {
		fTemp, err := excelize.OpenFile(tempFile)
		if err != nil {
			log.Fatal(err)
		}
		defer os.Remove(tempFile) // ลบไฟล์ชั่วคราวหลังจากรวมเสร็จ

		rows, err := fTemp.GetRows(sheetName)
		if err != nil {
			log.Fatal(err)
		}

		startRow := i*rowsPerWorker + 1
		for j, row := range rows {
			cell := fmt.Sprintf("A%d", startRow+j)
			if err := fFinal.SetSheetRow(sheetName, cell, &row); err != nil {
				log.Fatal(err)
			}
		}
	}

	if err := fFinal.SaveAs(finalFile); err != nil {
		log.Fatal(err)
	}

	// สิ้นสุดการวัดเวลา
	elapsed := time.Since(start)

	// แสดงผล elapsed time
	fmt.Printf("Elapsed time: %.2f seconds\n\n", elapsed.Seconds())
}

func generateExcelWithStreamWriter() {
	// เริ่มการวัดเวลา
	start := time.Now()

	// Create a new Excel file
	f := excelize.NewFile()

	// Create a new sheet
	sheetName := "Sheet1"
	index, err := f.NewSheet(sheetName)
	if err != nil {
		log.Fatal(err)
	}
	f.SetActiveSheet(index)

	// ใช้ StreamWriter เพื่อเขียนข้อมูล
	streamWriter, err := f.NewStreamWriter(sheetName)
	if err != nil {
		log.Fatal(err)
	}

	for i := 1; i <= 1000000; i++ {
		// เขียนข้อมูลทีละแถว
		row := []interface{}{
			fmt.Sprintf("Row %d", i),
			"Cell B",
		}
		cell, _ := excelize.CoordinatesToCellName(1, i)
		if err := streamWriter.SetRow(cell, row); err != nil {
			log.Fatal(err)
		}
	}

	// ปิด StreamWriter
	if err := streamWriter.Flush(); err != nil {
		log.Fatal(err)
	}

	// Save the file
	if err := f.SaveAs("LargeFile.xlsx"); err != nil {
		fmt.Println(err)
	}

	// สิ้นสุดการวัดเวลา
	elapsed := time.Since(start)

	// แสดงผล response status, body และ elapsed time
	fmt.Printf("Elapsed time: %.2f seconds\n\n", elapsed.Seconds())
}

func generateExcelConcurrent2() {
	maxWorker := 5
	batchSize := 1000

	// เริ่มการวัดเวลา
	start := time.Now()

	// Create a new Excel file
	f := excelize.NewFile()

	// Create a new sheet
	sheetName := "Sheet1"
	f.NewSheet(sheetName)

	// Create channel tasks
	tasks := make(chan []int, 1000000/batchSize)

	// Create wait group
	var wg sync.WaitGroup

	for i := 0; i < maxWorker; i++ {
		wg.Add(1)

		go func(w int) {
			defer wg.Done()

			for batch := range tasks {
				for _, task := range batch {
					cellA := fmt.Sprintf("A%d", task)
					cellB := fmt.Sprintf("B%d", task)
					f.SetCellValue(sheetName, cellA, fmt.Sprintf("Row %d", task))
					f.SetCellValue(sheetName, cellB, "Cell B")
				}
			}
		}(i)
	}

	// Send tasks to tasks channel in batches
	for i := 1; i <= 1000000; i += batchSize {
		end := i + batchSize - 1
		if end > 1000000 {
			end = 1000000
		}
		taskBatch := make([]int, 0, end-i+1)
		for j := i; j <= end; j++ {
			taskBatch = append(taskBatch, j)
		}
		tasks <- taskBatch
	}
	close(tasks)

	// Wait all go routine done
	wg.Wait()

	// Save the file
	if err := f.SaveAs("LargeFile.xlsx"); err != nil {
		fmt.Println(err)
	}

	// สิ้นสุดการวัดเวลา
	elapsed := time.Since(start)

	// แสดงผล response status, body และ elapsed time
	fmt.Printf("Elapsed time: %.2f seconds\n\n", elapsed.Seconds())
}

/* func generateExcelConcurrent() {
	maxWorker := 5

	// เริ่มการวัดเวลา
	start := time.Now()

	// Create a new Excel file
	f := excelize.NewFile()

	// Create a new sheet
	sheetName := "Sheet1"
	f.NewSheet(sheetName)

	// Create channel tasks
	tasks := make(chan int, 1000000)

	// Create wait group
	var wg sync.WaitGroup

	for i := 0; i < maxWorker; i++ {
		wg.Add(1)

		go func(w int) {
			defer wg.Done()

			for task := range tasks {
				// fmt.Printf("worker %d: %d\n", w, task)
				cellA := fmt.Sprintf("A%d", task)
				cellB := fmt.Sprintf("B%d", task)
				f.SetCellValue(sheetName, cellA, fmt.Sprintf("Row %d", task))
				f.SetCellValue(sheetName, cellB, "Cell B")
			}
		}(i)
	}

	// Send task to tasks channel
	for i := 0; i < 1000000; i++ {
		tasks <- i
	}
	close(tasks)

	// Wait all go routine done
	wg.Wait()

	// Save the file
	if err := f.SaveAs("LargeFile.xlsx"); err != nil {
		fmt.Println(err)
	}

	// สิ้นสุดการวัดเวลา
	elapsed := time.Since(start)

	// แสดงผล response status, body และ elapsed time
	fmt.Printf("Elapsed time: %.2f seconds\n\n", elapsed.Seconds())
} */

func generateExcel() {
	// เริ่มการวัดเวลา
	start := time.Now()

	// Create a new Excel file
	f := excelize.NewFile()

	// Create a new sheet
	sheetName := "Sheet1"
	f.NewSheet(sheetName)

	// Write 1,000,000 rows
	for i := 1; i <= 1000000; i++ {
		cellA := fmt.Sprintf("A%d", i)
		cellB := fmt.Sprintf("B%d", i)
		f.SetCellValue(sheetName, cellA, fmt.Sprintf("Row %d", i))
		f.SetCellValue(sheetName, cellB, "Cell B")
	}

	// Save the file
	if err := f.SaveAs("LargeFile.xlsx"); err != nil {
		fmt.Println(err)
	}

	// สิ้นสุดการวัดเวลา
	elapsed := time.Since(start)

	// แสดงผล response status, body และ elapsed time
	fmt.Printf("Elapsed time: %.2f seconds\n\n", elapsed.Seconds())
}

/* func readExcel() {
	// Open the Excel file
	f, err := excelize.OpenFile("LargeFile.xlsx")
	if err != nil {
		log.Fatal(err)
	}
	defer f.Close()

	// Get all sheet names
	sheetNames := f.GetSheetList()
	for _, sheetName := range sheetNames {
		fmt.Println("Reading sheet:", sheetName)

		// Get all rows in the sheet
		rows, err := f.GetRows(sheetName)
		if err != nil {
			log.Fatal(err)
		}

		// Print rows
		for _, row := range rows {
			for _, colCell := range row {
				fmt.Printf("%s\t\t", colCell)
			}
			fmt.Println()
		}
	}
} */
