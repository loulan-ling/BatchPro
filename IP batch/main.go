package main

import (
	"bufio"
	"fmt"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/sinlov/qqwry-golang/qqwry"
)

func main() {
	// 打开IP地址文件
	ipFile, err := os.Open("ip.txt")
	if err != nil {
		fmt.Println("无法打开IP地址文件:", err)
		return
	}
	defer ipFile.Close()

	// 创建 Excel 文件
	xlsx := excelize.NewFile()
	// 创建工作表
	index := xlsx.NewSheet("IP Results")
	// 设置表头
	xlsx.SetCellValue("IP Results", "A1", "IP Address")
	xlsx.SetCellValue("IP Results", "B1", "Country")
	xlsx.SetCellValue("IP Results", "C1", "Area")

	// 初始化 QQwry 数据库
	var datPath = "./qqwry.dat"
	qqwry.DatData.FilePath = datPath
	init := qqwry.DatData.InitDatFile()
	if v, ok := init.(error); ok {
		if v != nil {
			fmt.Printf("初始化 InitDatFile 出错: %s", v)
			return
		}
	}

	// 创建一个用于写入 Excel 的行号
	row := 2

	// 逐行读取IP地址文件，并查询并写入 Excel 文件
	scanner := bufio.NewScanner(ipFile)
	for scanner.Scan() {
		ip := scanner.Text()
		res := qqwry.NewQQwry().SearchByIPv4(ip)
		// 将查询结果写入 Excel 文件
		xlsx.SetCellValue("IP Results", fmt.Sprintf("A%d", row), ip)
		xlsx.SetCellValue("IP Results", fmt.Sprintf("B%d", row), res.Country)
		xlsx.SetCellValue("IP Results", fmt.Sprintf("C%d", row), res.Area)
		// 更新行号
		row++
	}

	// 将工作表设置为活动工作表
	xlsx.SetActiveSheet(index)

	// 将文件保存到指定路径
	err = xlsx.SaveAs("ip-results.xlsx")
	if err != nil {
		fmt.Println("保存 Excel 文件时出错:", err)
		return
	}

	fmt.Println("查询结果已成功写入到 ip-results.xlsx 文件.")
}

