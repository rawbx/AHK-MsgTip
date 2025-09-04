/**
 * @description show a tooltip for screen or GuiCtrl with custom style
 * @example <caption>Show screen tip:</caption>
 * tipBox := MsgTip()
 * tipBox.ShowScreenTip("message", timeout:=3000)
 * @example <caption>Hover GuiCtrl to show tip:</caption>
 * tipBox := MsgTip(timeout:=2000)
 * tmpGui := Gui()
 * btn := tmpGui.AddButton("w220 h50", "Hover it")
 * btn.AddTip("Hi! I am a btn")
 * tmpGui.Show()
 */
Class MsgTip {
    static HWND_DESKTOP := 0

    ; Tooltip styles
    static TTS_ALWAYSTIP := 0x01
    static TTS_NOPREFIX := 0x02
    static TTS_BALLOON := 0x40
    static TTS_CLOSE := 0x80

    ; Tooltip icons
    static TTI_NONE := 0
    static TTI_INFO := 1
    static TTI_WARNING := 2
    static TTI_ERROR := 3

    ; Tooltip flags
    static TTF_IDISHWND := 0x0001
    static TTF_CENTERTIP := 0x0002
    static TTF_RTLREADING := 0x0004
    static TTF_SUBCLASS := 0x0010
    static TTF_TRACK := 0x0020
    static TTF_ABSOLUTE := 0x0080
    static TTF_TRANSPARENT := 0x0100

    ; Tooltip message (WM_USER 0x400 + msg)
    static TTM_SETDELAYTIME := 0x403
    static TTM_TRACKACTIVATE := 0x411
    static TTM_TRACKPOSITION := 0x412
    static TTM_SETTIPBKCOLOR := 0x413
    static TTM_SETTIPTEXTCOLOR := 0x414
    static TTM_SETMAXTIPWIDTH := 0x418
    static TTM_ADDTOOL := 0x432
    static TTM_DELTOOL := 0x433
    static TTM_SETTITLE := 0x421
    static TTM_UPDATETIPTEXT := 0x439

    ; Tooltip message set delay time flag
    static TTDT_AUTOMATIC := 0
    static TTDT_RESHOW := 1
    static TTDT_AUTOPOP := 2
    static TTDT_INITIAL := 3


    static WS_EX_TOPMOST := 0x8
    static WS_EX_NOACTIVATE := 0x08000000
    static WS_POPUP := 0x080000000
    static WM_SETFONT := 0x30

    static SysDarkTheme := !RegRead(
        "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
        "AppsUseLightTheme", 1
    )

    /**
     * @param {Integer} timeout global duration, default is 3000ms
     * @param {Object} fontOpt font option, default is `{ name: "Microsoft YaHei UI", height: 30, weight: 400, italic: 0 }`
     * @param {Integer} bgColor hex rgb color: 0x212121
     * @param {Integer} textColor hex rgb color:0xffffff
     * @param {Integer} borderColor hex rgb color: 0xa0a0a0, **win11 require**
     * @param {Boolean} roundConner default is 1(enable), **win11 require**
     * @param {Boolean} balloonShape default is 0(rectangle shape)
     * @param {Boolean} multipleLine default is 1(enable)
     */
    __New(timeout := 3000, fontOpt?, bgColor?, textColor?, borderColor?, roundConner?, balloonShape?, multipleLine?) {
        this.hFont := 0
        this.hGui := MsgTip.HWND_DESKTOP
        this.balloonShape := balloonShape?? 0
        this.winExStyle := MsgTip.WS_EX_TOPMOST | MsgTip.WS_EX_NOACTIVATE
        this.ttStyle := MsgTip.TTS_ALWAYSTIP | MsgTip.TTS_NOPREFIX | (this.balloonShape? MsgTip.TTS_BALLOON : 0)
        this.ttFlags := MsgTip.TTF_TRACK | MsgTip.TTF_ABSOLUTE

        this.hTT := this.CreateWindowEx(this.hGui)
        this.sTi := this.CreateToolInfo(this.hGui)
        this.duration := timeout ; screen tip default duration
        this.SetCtrlTipDelay(this.duration) ; guiCtrl tip default duration
        this.SetAppearance(fontOpt?, bgColor?, textColor?, borderColor?, roundConner?, multipleLine?)

        if !Gui.Control.Prototype.HasProp("AddTip")
            Gui.Control.Prototype.DefineProp("AddTip", { Call: (ctrl, txt, title?) => this.SetCtrlTip(ctrl, txt,title?) })
    }

    __Delete() {
        if this.hFont
            DllCall("DeleteObject", "Ptr", this.hFont)
        if this.hTT
            DllCall("DestroyWindow", "Ptr", this.hTT)
    }

    ; https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-createwindowexw
    CreateWindowEx(hwnd) {
        local hTT := 0
        hTT := DllCall("CreateWindowEx",
            "UInt", this.winExStyle,                ; dwExStyle
            "Ptr", StrPtr("tooltips_class32"),      ; lpClassName
            "Ptr", 0,                               ; lpWindowName
            "UInt", this.ttStyle,                   ; dwStyle
            "Int", 0, "Int", 0, "Int", 0, "Int", 0, ; rect[x,y,nWidth,nHeight]
            "Ptr", hwnd,                            ; hWndParent
            "Ptr", 0, "Ptr", 0, "Ptr", 0,           ; hMenu, hInstance, lpParam
            "Ptr"
        )
        if !hTT
            throw Error("CreateWindowEx failed!")

        return hTT
    }

    ; https://learn.microsoft.com/windows/win32/api/commctrl/ns-commctrl-tttoolinfow
    CreateToolInfo(winHwnd) {
        static cbSize := 24 + A_PtrSize * 6
        local sTi := Buffer(cbSize, 0)                     ; toolinfo struct
        NumPut("UInt", cbSize, sTi, 0)                     ; cbSize
        ; NumPut("UInt", this.desktopTTFlag, sTi, 4)       ; uFlags
        ; NumPut("Ptr", winHwnd, sTi, 8)                   ; hwnd
        ; NumPut("UInt", 1, sTi, 16)                       ; uId
        NumPut("UInt", this.ttFlags, "UPtr", winHwnd, "UPtr", winHwnd, sTi, 4)
        return sTi
    }

    SetAppearance(fontOpt?, bgColor?, textColor?, borderColor?, roundConner?, multipleLine?) {
        if !isSet(bgColor) || !isSet(textColor) {
            DllCall("UxTheme\SetWindowTheme", "Ptr", this.hTT, "Ptr", MsgTip.SysDarkTheme ? StrPtr("DarkMode_Explorer") : StrPtr("Explorer"), "Ptr", 0)
        } else {
            DllCall("UxTheme\SetWindowTheme", "Ptr", this.hTT, "Ptr", 0, "Ptr", empty := 0)
            DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_SETTIPBKCOLOR, "Int", exchangeRbChannel(bgColor), "Int", 0)
            DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_SETTIPTEXTCOLOR, "Int", exchangeRbChannel(textColor), "Int", 0)
        }

        this.hFont := this.SetFont(fontOpt ?? { name: "Microsoft YaHei UI", height: 27, weight: 400, italic: 0 })

        ; only surport for win11 22000 and above
        if (VerCompare(A_OSVersion, "10.0.22000") > 0 && !this.balloonShape) {
            if isSet(borderColor)
                DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr", this.hTT, "UInt", DWMWA_BORDER_COLOR := 34, "Ptr*", &borderColor := exchangeRbChannel(borderColor), "UInt", 4)

            if roundConner ?? 1
                DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr", this.hTT, "UInt", DWMWA_WINDOW_CORNER_PREFERENCE := 33, "Ptr*", DWMWCP_ROUNDSMALL := 3, "UInt", 4)
        }

        if multipleLine ?? 1
            DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_SETMAXTIPWIDTH, "Ptr", 0, "Ptr", A_ScreenWidth)

        exchangeRbChannel(c) => ((c & 0x0000FF) << 16) | (c & 0x00FF00) | ((c & 0xFF0000) >> 16)
    }

    SetFont(fontOpt) {
        static sizeOfLogFontW := 92
        local sLf := Buffer(sizeOfLogFontW, 0)               ; lfStruct
        NumPut("Int", -fontOpt.height, sLf, 0)               ; lfHeight（negative as pt）
        NumPut("Int", fontOpt.weight, sLf, 16)               ; lfWeight
        NumPut("UChar", fontOpt.italic ? 1 : 0, sLf, 20)     ; lfItalic
        StrPut(fontOpt.name, sLf.Ptr + 28, 32, "UTF-16")     ; lfFaceName[32]
        local hFont := DllCall("gdi32\CreateFontIndirectW", "Ptr", sLf, "Ptr")
        DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.WM_SETFONT, "Ptr", hFont, "Int", 1)

        return hFont
    }

    /**
     * @param {GuiCtrlObj} guiCtrl
     * @param {String} txt tip text
     * @param {String} title tip title in balloonShape
     * @returns {GuiCtrlObj} guiCtrl of args
     */
    SetCtrlTip(guiCtrl, txt, title?) {
        if !txt
            return guiCtrl

        local ttFlags := MsgTip.TTF_SUBCLASS | MsgTip.TTF_IDISHWND
        NumPut("UInt", ttFlags, "UPtr", this.hGui, "UPtr", guiCtrl.Hwnd, this.sTi, 4) ; cbSize, uFlags, hwnd, uID
        DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_ADDTOOL, "Ptr", 0, "Ptr", this.sTi)
        if (IsSet(title) && this.balloonShape) {
            this.SetTitle(title)
        }
        this.UpdateTxt(txt)

        return guiCtrl
    }

    /**
     * @param {Integer} ms delay time(milliseconds), -1 is set it to default
     */
    SetCtrlTipDelay(ms) {
        DllCall("SendMessage",
            "Ptr", this.hTT,
            "UInt", MsgTip.TTM_SETDELAYTIME,
            "Ptr", MsgTip.TTDT_AUTOPOP,
            "Ptr", ms < -1 ? -1 : ms
        )
    }

    /**
     * @param {String} txt tip text
     * @param {String} title tip title in balloonShape
     * @param {Integer} timeout delay time(milliseconds)
     * @param {Integer} x tip coord x
     * @param {Integer} y tip coord x
     * @param {Integer} xOffset tip coord x offset
     * @param {Integer} yOffset tip coord y offset
     */
    ShowScreenTip(txt, title?, timeout?, x?, y?, xOffset := 25, yOffset := -15) {
        if !txt
            return

        DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_ADDTOOL, "Ptr", 0, "Ptr", this.sTi)

        local cursorPoint := Buffer(8, 0)
        DllCall("GetCursorPos", "Ptr", cursorPoint.Ptr)
        local mouseX := NumGet(cursorPoint, 0, "Int"), mouseY := NumGet(cursorPoint, 4, "Int")
        if (x ?? (x := mouseX + xOffset) && y ?? (y := mouseY - yOffset))
            DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_TRACKPOSITION, "Ptr", 0, "Ptr", (y << 16) | (x & 0xFFFF))

        if (IsSet(title) && this.balloonShape) {
            this.SetTitle(title)
        }

        this.UpdateTxt(txt)

        display(1) ; 1 is show, 0 is hide
        SetTimer((*) => display(0), -Abs(timeout ?? this.duration))

        display(vis) {
            DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_TRACKACTIVATE, "Int", vis, "Ptr", this.sTi)
        }
    }

    SetTitle(txt, iconType := MsgTip.TTI_INFO) {
        if txt
            DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_SETTITLE, "Ptr", iconType, "Ptr", StrPtr(txt))
    }

    UpdateTxt(txt) {
        NumPut("UPtr", StrPtr(txt), this.sTi, 24 + (A_PtrSize * 3))  ; lpszText
        DllCall("SendMessage", "Ptr", this.hTT, "UInt", MsgTip.TTM_UPDATETIPTEXT, "Ptr", 0, "Ptr", this.sTi)
    }
}
