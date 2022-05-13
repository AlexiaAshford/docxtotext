package main

import (
	"baliance.com/gooxml/document"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"os"
	"os/exec"
	"path"
	"path/filepath"
	"strings"
)

type Config struct {
	DocToDocx   bool
	DelDocFile  bool
	DelDocxFile bool
}
type ConfigClass struct {
	ConfigFile      string
	FileInformation []byte
	FileNameList    []string
	FileStruct      Config
}

func (is *ConfigClass) SaveConfig() {
	if err := ioutil.WriteFile("./config.json", is.FileInformation, 0777); err != nil {
		log.Fatalf("error writing file: %s", err)
	}
}
func (is *ConfigClass) load() {
	if data, err := ioutil.ReadFile("./config.json"); err != nil {
		is.FileInformation = data
	}
}

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

func CmdPythonSaveDocx() {
	if _, err := exec.Command("python", []string{"run.py"}...).Output(); err == nil {
		fmt.Println("doc转换成功")
	} else {
		fmt.Println(err)
	}
}

func (is *ConfigClass) switchFileName() bool {
	if NameList := is.FileNameList; NameList != nil || len(NameList) != 0 {
		for index, file := range is.FileNameList {
			fileName := filepath.Base(file)
			switch path.Ext(fileName) {
			case ".docx":
				if docxContent := getDocxInformation(fileName); docxContent != "" {
					saveFile(strings.Replace(fileName, ".docx", ".txt", -1), docxContent)
					if is.FileStruct.DelDocxFile {
						if err := os.Remove(fileName); err != nil {
							log.Println(err)
						}
					}
				} else {
					fmt.Println("文件" + fileName + "处理失败")
				}
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

func init() {
	if _, err := os.Stat("./TextFile"); err != nil {
		if err = os.Mkdir("./TextFile", 0777); err != nil {
			log.Fatalf("error creating directory: %s", err)
		}
	} else {
		fmt.Println("TextFile 目录已存在, 不再创建")
	}
}

func initConfig() *ConfigClass {
	Vars := ConfigClass{ConfigFile: "./config.json"}
	if _, err := os.Stat(Vars.ConfigFile); err != nil {
		if config, ok := json.MarshalIndent(&Config{}, "", "   "); ok == nil {
			Vars.FileInformation = config
			Vars.SaveConfig()
		} else {
			log.Fatalf("error marshal config: %s", ok)
		}
	} else {
		if data, err := ioutil.ReadFile("./config.json"); err == nil {
			Vars.FileInformation = data
		} else {
			log.Fatalf("error reading file: %s", err)
		}
	}
	Vars.FileNameList = FileNameList() // 获取当前目录下所有文件名
	return &Vars
}
func main() {
	Vars := initConfig()
	if err := json.Unmarshal(Vars.FileInformation, &Vars.FileStruct); err == nil {
		if Vars.FileStruct.DocToDocx {
			CmdPythonSaveDocx() // 调用python脚本转换doc为docx
		}
		if !Vars.switchFileName() {
			log.Println("文件列表获取失败或没有查找到doc docx 文档")
		}
	}
}
