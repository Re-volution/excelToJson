package main

import (
	"fmt"
	"os"
)

var folder = "./excel"
var outFolder = "./json"

func readFile() {
	os.RemoveAll(outFolder)
	files, err := os.ReadDir(folder)
	if err != nil {
		fmt.Println("读取目录文件错误:", err)
		return
	}
	for _, file := range files {
		if file.IsDir() {

		} else {
			excel(folder + "/" + file.Name())
		}
	}
}

func write(file string, data string) {
	if !isPathExist(outFolder) {
		os.Mkdir(outFolder, 0777)
	}
	obj, err := os.Create(file)
	if err != nil {
		fmt.Println(err)
	}
	obj.WriteString(data)
	obj.Close()
}

func isPathExist(path string) bool {
	if _, err := os.Stat(path); err == nil {
		return true
	} else if os.IsNotExist(err) {
		return false
	}
	return false
}
