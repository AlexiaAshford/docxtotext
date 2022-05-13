package config

import (
	"fmt"
	"os/exec"
)

func CmdPythonSaveDocx() {
	if _, err := exec.Command("python", []string{"run.py"}...).Output(); err == nil {
		fmt.Println("doc转换成功")
	} else {
		fmt.Println(err)
	}
}
