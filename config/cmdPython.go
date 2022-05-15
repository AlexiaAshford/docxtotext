package config

import (
	"fmt"
	"os/exec"
)

func CmdPythonSaveDocx(arg []string) {
	if _, err := exec.Command("python", arg...).Output(); err == nil {
		fmt.Println("doc转换成功")
	} else {
		fmt.Println(err)
	}
}
