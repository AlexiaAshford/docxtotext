package main

import (
	"baliance.com/gooxml/document"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"path"
	"path/filepath"
	"strings"
)

func FileNameList() []string {
	var files []string
	if dir, err := os.Getwd(); err == nil {
		if ok := filepath.Walk(dir, func(path string, info os.FileInfo, err error) error {
			files = append(files, path)
			return nil
		}); ok == nil {
			return files
		}
	} else {
		fmt.Println(err)
	}
	return nil
}

func MkdirTextFile(path string) {
	if _, err := os.Stat(path); err != nil {
		if err = os.Mkdir(path, 0777); err != nil {
			log.Fatalf("error creating directory: %s", err)
		}
	} else {
		fmt.Println("TextFile 目录已存在, 不再创建")
	}
}
func switchFileName() bool {
	if NameList := FileNameList(); NameList != nil || len(NameList) != 0 {
		for index, file := range FileNameList() {
			fileName := filepath.Base(file)
			switch path.Ext(fileName) {
			case ".docx":
				saveFile(strings.Replace(fileName, ".docx", ".txt", -1), getDocxInformation(fileName))
			case ".doc":
				saveFile(strings.Replace(fileName, ".doc", ".txt", -1), getDocxInformation(fileName))
			default:
				fmt.Println("No:", index, fileName, "不是docx文件，跳过处理！")
			}
		}
		return true
	}
	return false
}

func saveFile(fileName, content string) {
	content = "　　" + strings.Replace(content, "\n　　\n", "\n", -1)
	if ok := ioutil.WriteFile("TextFile/"+fileName, []byte(content), 0644); ok != nil {
		log.Fatalf("error writing file: %s", ok)
	} else {
		fmt.Println("文件" + fileName + "处理完毕， 已保存在 TextFile 目录下")
	}
}

func getDocxInformation(fileName string) string {
	var content string
	//doc.Paragraphs() 得到包含文档所有的段落的切片
	if doc, err := document.Open(fileName); err == nil {
		for _, para := range doc.Paragraphs() {
			//run为每个段落相同格式的文字组成的片段
			for _, run := range para.Runs() {
				content += run.Text()
			}
			content += "\n　　"
		}
	} else {
		log.Fatalf("error opening document: %s", err)
	}
	return content
}

func main() {
	MkdirTextFile("./TextFile")
	if !switchFileName() {
		log.Println("文件列表获取失败或没有查找到doc docx 文档")
	}
}
