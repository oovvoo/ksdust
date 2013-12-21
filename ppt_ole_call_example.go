package main

import (
	"fmt"
	"syscall"
	"unsafe"
	"os"
	"bufio"
	"strings"
)

const noteBufferLen = 1024

func main() {
	h, err := syscall.LoadLibrary("PowerPointWrapper.dll")
	if (err != nil) {

	}

	defer syscall.FreeLibrary(h)
	procInitialize, err := syscall.GetProcAddress(h, "Initialize")
	if err != nil {
		fmt.Println(err.Error())
	}
	procPresentationCurrentSlideNote, err := syscall.GetProcAddress(h, "PresentationCurrentSlideNote")
	if err != nil {
		fmt.Println(err.Error())
	}
	procUninitialize, err := syscall.GetProcAddress(h, "Uninitialize")
	if err != nil {
		fmt.Println(err.Error())
	}

	var noteBuffer [noteBufferLen]uint8

	_, _, err = syscall.Syscall(uintptr(procInitialize), 0, 0, 0, 0)
	
	for i := 0; i <3; i++ {
		noteLen, _, err := syscall.Syscall(uintptr(procPresentationCurrentSlideNote), 2, uintptr(unsafe.Pointer(&noteBuffer)), uintptr(noteBufferLen), 0)
		if err != 0  {
			fmt.Println(err)
			continue
		}

		note := string(noteBuffer[:int(noteLen)])
		replacedNote := strings.Replace(note, "\r", "\n", -1)
		//writeLines([]string{note}, "b.txt")

		fmt.Printf("note: `%s`nstrlen: %d\n", replacedNote, len(note))
	}
	_, _, err = syscall.Syscall(uintptr(procUninitialize), 0, 0, 0, 0)

	
}

func writeLines(lines []string, path string) error {
  file, err := os.Create(path)
  if err != nil {
    return err
  }
  defer file.Close()

  w := bufio.NewWriter(file)
  for _, line := range lines {
  	//fmt.Printf("%d\n", '\r')
  	for _, c := range line {
		fmt.Printf("%d\n", c)
	}
    fmt.Fprintln(w, strings.Replace(line, "\r", "\n", -1))
  }
  return w.Flush()
}