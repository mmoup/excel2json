package main

import (
	"fmt"
	mpath "mmoup/path"
	"os"
	"path"
	"strconv"
	"strings"

	"github.com/tealeg/xlsx"
)

type CellInfo struct {
	Dtype string
	Name  string
	Desc  string
}

func main() {
	fmt.Println(os.Args)
	excelFolder := mpath.GetCurrentDirectory()
	jsonFolder := excelFolder
	portrait := false

	if len(os.Args) > 1 {
		if os.Args[1] == "portrait" {
			portrait = true
		} else {
			excelFolder = os.Args[1]
			jsonFolder = path.Dir(excelFolder)
		}
	}

	if len(os.Args) > 2 {
		if portrait {
			excelFolder = os.Args[2]
			jsonFolder = path.Dir(excelFolder)
		} else {
			jsonFolder = os.Args[2]
		}
	}

	if len(os.Args) > 3 {
		if portrait {
			jsonFolder = os.Args[3]
		}
	}

	if strings.HasSuffix(excelFolder, ".xlsx") {
		jsonFile := jsonFolder + "/" + strings.TrimSuffix(path.Base(excelFolder), ".xlsx") + ".json"
		if portrait {
			parseExcelPortrait(excelFolder, jsonFile)
		} else {
			parseExcel(excelFolder, jsonFile)
		}
	} else {
		for _, file := range mpath.GetFilelist(excelFolder) {
			if strings.HasSuffix(file.Name(), ".xlsx") && !strings.HasPrefix(file.Name(), "~$") {
				excelFile := excelFolder + "/" + file.Name()
				jsonFile := jsonFolder + "/" + strings.TrimSuffix(file.Name(), ".xlsx") + ".json"
				if portrait {
					parseExcelPortrait(excelFile, jsonFile)
				} else {
					parseExcel(excelFile, jsonFile)
				}

			}
		}
	}

	fmt.Println("转换完成")
}

func parseExcel(excelFile, jsonFile string) bool {
	fmt.Println("转换", excelFile, "到", jsonFile)
	xlFile, err := xlsx.OpenFile(excelFile)
	if err != nil {
		fmt.Println(err.Error())
	}

	f, err := os.Create(jsonFile)
	if err != nil {
		panic(err)
	}
	defer f.Close()

	f.WriteString("{")

	firstSheet := true
	for _, sheet := range xlFile.Sheets {
		//fmt.Println(sheet.Name)
		if firstSheet {
			f.WriteString("\r\n\t\"")
			firstSheet = false
		} else {
			f.WriteString(",\r\n\t\"")
		}
		f.WriteString(sheet.Name)
		f.WriteString("\" : [")

		cellType := make([]string, 0, len(sheet.Cols))
		cellName := make([]string, 0, len(sheet.Cols))

		for row := 0; row < len(sheet.Rows); row++ {
			//跳过空行

			if len(sheet.Rows[row].Cells) == 0 || strings.TrimSpace(sheet.Rows[row].Cells[0].String()) == "" {
				break
			}

			if row == 3 {
				f.WriteString("\r\n\t\t{")
			} else if row > 3 {
				f.WriteString(",\r\n\t\t{")
			}

			for col := 0; col < len(sheet.Rows[row].Cells); col++ {
				//如果column没有定义数据类型，不处理此column的数据
				if row > 0 && col >= len(cellType) {
					break
				}

				cellValue := strings.TrimSpace(sheet.Rows[row].Cells[col].String())
				if row == 0 {
					if cellValue == "" {
						break
					}
					cellType = append(cellType, cellValue)
				} else if row == 1 {
					cellName = append(cellName, cellValue)
				} else if row == 2 {
				} else {
					if len(cellType[col]) > 0 && strings.Contains("int|bool|string|float", cellType[col]) {
						if col == 0 {
							f.WriteString("\r\n\t\t\t\"")
						} else {
							f.WriteString(",\r\n\t\t\t\"")
						}
					}

					switch cellType[col] {
					case "int":
						f.WriteString(cellName[col])
						f.WriteString("\" : ")
						cellValueInt, err := sheet.Rows[row].Cells[col].Int()
						if err != nil {
							fmt.Errorf(err.Error())
						}
						cellValue := strconv.Itoa(cellValueInt)
						f.WriteString(cellValue)
					case "string":
						f.WriteString(cellName[col])
						f.WriteString("\" : \"")
						f.WriteString(strings.Replace(cellValue, "\n", "\\n", -1))
						f.WriteString("\"")
					case "bool":
						f.WriteString(cellName[col])
						f.WriteString("\" : ")
						if cellValue == "true" || cellValue == "1" || cellValue == "TRUE" {
							f.WriteString("true")
						} else {
							f.WriteString("false")
						}
					case "float":
						f.WriteString(cellName[col])
						f.WriteString("\" : ")
						cellValueFloat, err := sheet.Rows[row].Cells[col].Float()
						if err != nil {
							fmt.Errorf(err.Error())
						}
						cellValue = strconv.FormatFloat(cellValueFloat, 'f', -1, 32)
						f.WriteString(cellValue)
					}
				}
			}
			if row >= 3 {
				f.WriteString("\r\n\t\t}")
			}
		}
		f.WriteString("\r\n\t]")
	}
	f.WriteString("\r\n}")
	return true
}

func parseExcelPortrait(excelFile, jsonFile string) bool {
	fmt.Println("转换", excelFile, "到", jsonFile, "竖排格式")
	xlFile, err := xlsx.OpenFile(excelFile)
	if err != nil {
		fmt.Println(err.Error())
	}

	f, err := os.Create(jsonFile)
	if err != nil {
		panic(err)
	}
	defer f.Close()

	f.WriteString("{\r\n")

	firstSheet := true
	for _, sheet := range xlFile.Sheets {
		//fmt.Println(sheet.Name)
		if firstSheet {
			f.WriteString("\r\n\t\"")
			firstSheet = false
		} else {
			f.WriteString(",\r\n\t\"")
		}
		f.WriteString(sheet.Name)
		f.WriteString("\" : {")

		firstRow := true
		for row := 0; row < len(sheet.Rows); row++ {
			cellType := ""
			for col := 0; col < len(sheet.Rows[row].Cells); col++ {
				//如果column没有定义数据类型，不处理此column的数据
				cellValue := strings.TrimSpace(sheet.Rows[row].Cells[col].String())
				//值类型
				if col == 0 {
					if cellValue == "" {
						break
					}
					if !strings.Contains("int|bool|string|float", cellValue) {
						break
					}
					cellType = cellValue
				} else if col == 1 { //名字
					if firstRow {
						f.WriteString("\r\n\t\t\"" + cellValue + "\" : ")
						firstRow = false
					} else {
						f.WriteString(",\r\n\t\t\"" + cellValue + "\" : ")
					}
				} else if col == 2 { //描述
				} else if col == 3 { //值
					switch cellType {
					case "int":
						cellValueInt, err := sheet.Rows[row].Cells[col].Int()
						if err != nil {
							fmt.Errorf(err.Error())
						}
						cellValue = strconv.Itoa(cellValueInt)
						f.WriteString(cellValue)
					case "string":
						f.WriteString("\"" + cellValue + "\"")
					case "bool":
						if cellValue == "true" || cellValue == "1" || cellValue == "TRUE" {
							f.WriteString("true")
						} else {
							f.WriteString("false")
						}
					case "float":
						cellValueFloat, err := sheet.Rows[row].Cells[col].Float()
						if err != nil {
							fmt.Errorf(err.Error())
						}
						cellValue = strconv.FormatFloat(cellValueFloat, 'f', -1, 32)
						f.WriteString(cellValue)
					}
				}
			}
		}
		f.WriteString("\r\n\t}")
	}
	f.WriteString("\r\n}")
	return true
}
