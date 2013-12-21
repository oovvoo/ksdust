// pptctrl 包裝了 PowerPointWrapper.dll 的函數呼叫。此
// package 內的「所有」方法呼叫都必須在同一個 OS thread 內進行。
package pptctrl

import (
	"fmt"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"syscall"
	"unsafe"
)

type Request struct {
	Data     []byte
	Response chan string
}

const _NoteBufferLength = 4096
const _NameBufferLength = 4096

var (
	eventLoopStarted    bool        = false
	startEventLoopMutex *sync.Mutex = new(sync.Mutex)
	eventChannel                    = make(chan *Request)

	noteBuffer [_NoteBufferLength]uint8
	nameBuffer [_NameBufferLength]uint8
)

// proc handles
var (
	dll syscall.Handle

	procInitialize   uintptr
	procUninitialize uintptr

	procPresentationCurrentSlideIndex uintptr
	procPresentationTotalSlidesCount  uintptr

	procPresentationPreviousSlide uintptr
	procPresentationNextSlide     uintptr

	procPresentationCurrentSlideName uintptr
	procPresentationCurrentSlideNote uintptr

	procRefreshPresentationSlidesThumbnail uintptr
)

func init() {
	dll, err := syscall.LoadLibrary("PowerPointWrapper.dll")
	if err != nil {
		panic(err)
	}

	procInitialize, err = syscall.GetProcAddress(dll, "Initialize")
	if err != nil {
		panic(err)
	}

	procUninitialize, err = syscall.GetProcAddress(dll, "Uninitialize")
	if err != nil {
		panic(err)
	}

	procPresentationCurrentSlideIndex, err = syscall.GetProcAddress(dll, "PresentationCurrentSlideIndex")
	if err != nil {
		panic(err)
	}

	procPresentationTotalSlidesCount, err = syscall.GetProcAddress(dll, "PresentationTotalSlidesCount")
	if err != nil {
		panic(err)
	}

	procPresentationPreviousSlide, err = syscall.GetProcAddress(dll, "PresentationPreviousSlide")
	if err != nil {
		panic(err)
	}

	procPresentationNextSlide, err = syscall.GetProcAddress(dll, "PresentationNextSlide")
	if err != nil {
		panic(err)
	}

	procPresentationCurrentSlideName, err = syscall.GetProcAddress(dll, "PresentationCurrentSlideName")
	if err != nil {
		panic(err)
	}

	procPresentationCurrentSlideNote, err = syscall.GetProcAddress(dll, "PresentationCurrentSlideNote")
	if err != nil {
		panic(err)
	}

	procRefreshPresentationSlidesThumbnail, err = syscall.GetProcAddress(dll, "RefreshPresentationSlidesThumbnail")
	if err != nil {
		panic(err)
	}
}

// 初始化 COM 函式庫，在進行任何操作前必須先呼叫此方法。
func initialize() {
	_, _, err := syscall.Syscall(uintptr(procInitialize), 0, 0, 0, 0)
	if err != 0 {
		// nothing can be done without the successfully call of Initialize()
		panic(err)
	}
}

// 釋放 COM 函式庫的資源，不再使用這個 package 的功能時可呼叫此方法釋放資源。
func uninitialize() {
	_, _, err := syscall.Syscall(uintptr(procUninitialize), 0, 0, 0, 0)
	if err != 0 {
		fmt.Println("Uninitialize error:", err)
	}
}

// 傳回現在正在放映的投影片 index，若發生任何錯誤，傳回 -1
func presentationCurrentSlideIndex() int32 {
	r0, _, err := syscall.Syscall(uintptr(procPresentationCurrentSlideIndex), 0, 0, 0, 0)
	if err != 0 {
		fmt.Println("PresentationCurrentSlideIndex error:", err)
	}

	return int32(r0)
}

// 傳回現在正在放映的簡報所含的投影片數量，若發生任何錯誤，傳回 -1
func presentationTotalSlidesCount() int32 {
	r0, _, err := syscall.Syscall(uintptr(procPresentationTotalSlidesCount), 0, 0, 0, 0)
	if err != 0 {
		fmt.Println("PresentationTotalSlidesCount error:", err)
	}

	return int32(r0)
}

// 控制簡報：前往上一張投影片
func presentationPreviousSlide() {
	_, _, err := syscall.Syscall(uintptr(procPresentationPreviousSlide), 0, 0, 0, 0)
	if err != 0 {
		fmt.Println("PresentationPreviousSlide error:", err)
	}
}

// 控制簡報：前往下一張投影片
func presentationNextSlide() {
	_, _, err := syscall.Syscall(uintptr(procPresentationNextSlide), 0, 0, 0, 0)
	if err != 0 {
		fmt.Println("PresentationNextSlide error:", err)
	}
}

// 傳回目前放映的投影片的名稱
func presentationCurrentSlideName() string {
	nameLen, _, err := syscall.Syscall(uintptr(procPresentationCurrentSlideName), 2, uintptr(unsafe.Pointer(&nameBuffer)), uintptr(_NameBufferLength), 0)
	if err != 0 {
		fmt.Println("PresentationCurrentSlideName error:", err)
	}

	if nameLen < 1 {
		return "(未取得投影片名稱)"
	}

	name := string(nameBuffer[:int(nameLen)])
	return strings.Replace(name, "\r", "<br />", -1)
}

// 傳回目前放映的投影片的備忘稿
func presentationCurrentSlideNote() string {
	noteLen, _, err := syscall.Syscall(uintptr(procPresentationCurrentSlideNote), 2, uintptr(unsafe.Pointer(&noteBuffer)), uintptr(_NoteBufferLength), 0)
	if err != 0 {
		fmt.Println("PresentationCurrentSlideNote error:", err)
	}

	if noteLen < 1 {
		return "(未取得備忘稿)"
	}

	note := string(noteBuffer[:int(noteLen)])
	return strings.Replace(note, "\r", "<br />", -1)
}

// deletes all file in a foler, and create thumbnails of active slideshow.
//
// note: failed execution of this function will NOT recover any files previous existed in the folder.
func refreshPresentationSlidesThumbnail() {
	_, _, err := syscall.Syscall(uintptr(procRefreshPresentationSlidesThumbnail), 0, 0, 0, 0)
	if err != 0 {
		fmt.Println("RefreshPresentationSlidesThumbnail error:", err)
	}
}

func SendRequest(request *Request) {
	if !eventLoopStarted {
		panic("pptctrl.StartEventLoop() hasn't been called")
	}

	eventChannel <- request
}

func StartEventLoop() {
	if eventLoopStarted {
		return
	}

	startEventLoopMutex.Lock()

	if eventLoopStarted {
		return
	}

	go runEventLoop()
	eventLoopStarted = true
	startEventLoopMutex.Unlock()
}

// this function will locked to a OS thread and never returns, so it needs a channel to transfer messages
func runEventLoop() {
	runtime.LockOSThread()
	initialize()

	defer func() {
		startEventLoopMutex.Lock()

		//close(eventChannel)
		uninitialize()
		runtime.UnlockOSThread()
		eventLoopStarted = false

		startEventLoopMutex.Unlock()
	}()

	for {
		select {
		case c := <-eventChannel:
			switch c.Data[1] {
			case 'l': // 要求下一張投影片的縮圖網址
				nextSlideIndex := presentationCurrentSlideIndex() + 1
				c.Response <- (`/presthumbnail/` + strconv.Itoa(int64(nextSlideIndex)))
			case 'p': // 切換下一張投影片 & 要求目前投影片的備忘稿
				presentationPreviousSlide()
				c.Response <- presentationCurrentSlideNote()
			case 'n': // 切換上一張投影片 & 要求目前投影片的備忘稿
				presentationNextSlide()
				c.Response <- presentationCurrentSlideNote()
			case 'r': // 要求目前投影片的備忘稿
				c.Response <- presentationCurrentSlideNote()
			case 'h': // 重新整理投影片縮圖
				refreshPresentationSlidesThumbnail()
				c.Response <- ""
			}
		}
	}
}
