package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/360EntSecGroup-Skylar/excelize"
	"os"
	"path"
	"strconv"
	"strings"
	"time"
)

var sheetN *string
var checkStr = "dwa645d3k.,xf"

func init() {
	sheetN = flag.String("s", "data", "excel表的页签名称")
	flag.Parse() //解析输入的参数
}

func excel(file string) {
	start := time.Now()

	xlsx, err := excelize.OpenFile(file)
	if err != nil {
		fmt.Println(err)
		return
	}

	rows := xlsx.GetRows(*sheetN)
	var names []string
	var typs []string
	list := make([]map[string]interface{}, 0, len(rows))

	for i, row := range rows {
		hang := make(map[string]interface{}, len(row))

		for num, text := range row {
			if i == 0 {
				names = append(names, text)
			} else if i == 1 {
				typs = append(typs, text)
			} else if num < len(names) && num < len(typs) {
				hang[names[num]] = getData(typs[num], text)
			} else {
				fmt.Println("超长未定义的数据:", text)
			}

		}

		if i >= 2 {
			list = append(list, hang)
		}
	}

	//格式化为json
	data, err := json.Marshal(list)
	if err != nil {
		fmt.Printf("json.marshal failed,err:", err)
		return
	}

	fileSuffix := path.Ext(file) //获取文件后缀
	finame := path.Base(file)
	filenameOnly := strings.TrimSuffix(finame, fileSuffix) //获取文件名
	json_dir := outFolder + "/" + filenameOnly + ".json"

	if strings.Contains(string(data), checkStr) {
		fmt.Println("文件解析失败，请检查文件:", finame)
		os.Exit(1)
	}

	write(json_dir, string(data))
	//运行时间
	cost := time.Since(start)
	fmt.Println(file, "is success!", cost)
}

func getData(typ, text string) interface{} {
	count := strings.Count(typ, "[]")
	sep := ""
	switch count {
	case 1:
		sep = ","
	case 2:
		sep = ";"
	default:
	}
	if sep != "" { //数组
		strs := strings.Split(text, sep)
		var rd []interface{}
		for _, v := range strs {
			rd = append(rd, getData(typ[2:], v))
		}
		return rd
	}

	switch typ {
	case "int":
		//判断并转为int
		t, err := strconv.Atoi(text)
		if err == nil {
			return t
		} else {
			fmt.Println("int参数不对，无法转换成数字:", text)
			return checkStr
		}
	case "float":
		//判断并转为float64
		f, err := strconv.ParseFloat(text, 64)
		if err == nil {
			return f
		} else {
			fmt.Println("float参数不对，无法转换成小数:", text)
			return checkStr
		}
	case "string":
		return text
	default:
		fmt.Println("typ 不存在:", typ)
		return checkStr
	}
}
