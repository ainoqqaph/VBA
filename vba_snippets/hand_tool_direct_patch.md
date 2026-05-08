# HAND TOOL 直接補丁

如果原始資料本來沒有 `HAND TOOL`，請用這段補丁。

建議先在 VBA 模組最上方，`Private Info As Variant` 下方加入：

```vb
Private Const DEFAULT_INV_TITLE As String = "HAND TOOL"
```

然後在 `CollectDataTinv` 內，這一行之前：

```vb
PO = Array(Title, PoNo, isPOinFront)
```

加入：

```vb
If Len(Trim(CStr(Title))) = 0 Then
    Title = DEFAULT_INV_TITLE
End If
```

如果這個客戶所有商品都一定要歸在 `HAND TOOL`，可以改成強制覆蓋：

```vb
Title = DEFAULT_INV_TITLE
```
