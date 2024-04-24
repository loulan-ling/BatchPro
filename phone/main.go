package main

import (
	"bufio"
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/xluohome/phonedata"
)

func main() {
	err := os.Setenv("PHONE_DATA_DIR", "./")
	if err != nil {
		fmt.Println("设置环境变量 PHONE_DATA_DIR 失败:", err)
		return
	}
	// 打开 iphone.txt 文件
	file, err := os.Open("phone.txt")
	if err != nil {
		fmt.Println("无法打开文件:", err)
		return
	}
	defer file.Close()

	// 创建 Excel 文件
	xlsx := excelize.NewFile()
	// 创建工作表
	index := xlsx.NewSheet("Phone Results")
	// 设置表头
	xlsx.SetCellValue("Phone Results", "A1", "Phone Number")
	xlsx.SetCellValue("Phone Results", "B1", "Area Zone")
	xlsx.SetCellValue("Phone Results", "C1", "Card Type")
	xlsx.SetCellValue("Phone Results", "D1", "City")
	xlsx.SetCellValue("Phone Results", "E1", "Zip Code")
	xlsx.SetCellValue("Phone Results", "F1", "Province")

	// 逐行读取 iphone.txt 文件，并查询并写入 Excel 文件
	scanner := bufio.NewScanner(file)
	row := 2 // 行号
	for scanner.Scan() {
		phone := scanner.Text()
		pr, err := phonedata.Find(phone)
		if err != nil {
			fmt.Printf("查询手机号码 %s 时出错: %s\n", phone, err)
			continue
		}
		// 写入查询结果到 Excel 文件
		xlsx.SetCellValue("Phone Results", fmt.Sprintf("A%d", row), phone)
		xlsx.SetCellValue("Phone Results", fmt.Sprintf("B%d", row), pr.AreaZone)
		xlsx.SetCellValue("Phone Results", fmt.Sprintf("C%d", row), pr.CardType)
		xlsx.SetCellValue("Phone Results", fmt.Sprintf("D%d", row), pr.City)
		xlsx.SetCellValue("Phone Results", fmt.Sprintf("E%d", row), pr.ZipCode)
		xlsx.SetCellValue("Phone Results", fmt.Sprintf("F%d", row), pr.Province)
		// 更新行号
		row++
	}

	// 将工作表设置为活动工作表
	xlsx.SetActiveSheet(index)

	// 将文件保存到指定路径
	err = xlsx.SaveAs("phone-results.xlsx")
	if err != nil {
		fmt.Println("保存 Excel 文件时出错:", err)
		return
	}

	fmt.Println("查询结果已成功写入到 phone-results.xlsx 文件.")
}

