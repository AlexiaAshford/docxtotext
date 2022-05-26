package config

import (
	"encoding/json"
	"io/ioutil"
	"log"
	"os"
)

type Config struct {
	DocToDocx   bool
	DelDocFile  bool
	DelDocxFile bool
}

type ClassConfig struct {
	ConfigFileName  string
	TextFileName    string
	DocxFileName    string
	FileInformation []byte
	FileNameList    []string
	FileStruct      Config
	DelFileList     []string
}

func (is *ClassConfig) SaveConfig() {
	if err := ioutil.WriteFile(is.ConfigFileName, is.FileInformation, 0777); err != nil {
		log.Fatalf("error writing file: %s", err)
	}
}
func (is *ClassConfig) load() {
	if data, err := ioutil.ReadFile(is.ConfigFileName); err != nil {
		is.FileInformation = data
	}
}

func InitConfig() *ClassConfig {
	Vars := ClassConfig{
		ConfigFileName: "./config.json",
		TextFileName:   "./TextFile/",
		DocxFileName:   "./DocxFile/",
	}
	if _, err := os.Stat(Vars.ConfigFileName); err != nil {
		if configs, ok := json.MarshalIndent(&Vars.FileStruct, "", "   "); ok == nil {
			Vars.FileInformation = configs
			Vars.SaveConfig()
		} else {
			log.Fatalf("error marshal config: %s", ok)
		}
	} else {
		if data, err := ioutil.ReadFile(Vars.ConfigFileName); err == nil {
			Vars.FileInformation = data
		} else {
			log.Fatalf("error reading file: %s", err)
		}
	}
	Vars.FileNameList = FileNameList(Vars.DocxFileName) // 获取当前目录下所有文件名
	return &Vars
}
