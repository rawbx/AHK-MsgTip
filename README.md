# AHK-MsgTip

The default AHK's tooltip is ugly and cannot be customized, nor can it be used as a tip for GUI controls, so I wrote this library.

Some colors and styles that can be set, such as: fontOption, backgroundColor, textColor, borderColor, roundedConner, shape, etc.
If you don't want to customize, MsgTip will follow the system theme(dark or light).

## Usage

Put `MsgTip.ahk` to you lib folder.

If you want to use it as screen tootip.

```autoit
#Include <MsgTip>

tip := MsgTip()
tip.ShowScreenTip("message", , timeout:=3000)
```

Or show a tip to your gui control.

```autoit
#Include <MsgTip>

tip := MsgTip(timeout:=2000)
tmpGui := Gui()
btn := tmpGui.AddButton("w220 h50", "Hover it")
btn.AddTip("Hi! I am a btn")
tmpGui.Show()
```
