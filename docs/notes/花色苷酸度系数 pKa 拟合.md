### 花色苷酸度系数 pKa 拟合

#### 1. 花色苷质子转移理论模型

花色苷的多个酚羟基都可以解离释放氢离子，因此花色苷呈多元弱酸特性，这里以 $\mathrm{A}$ 、$\mathrm{A^-}$、$\mathrm{A^{2-}}$ 为例进行介绍之子转移/酸碱反应：

$$
\begin{equation}\begin{aligned}
&\mathrm{AH^+} \rightleftharpoons \mathrm{A} + \mathrm{H^+} &&
K_\mathrm{a1} = \frac{[\mathrm{A}][\mathrm{H^+}]}{[\mathrm{AH^+}]}\\
&\mathrm{A} \rightleftharpoons \mathrm{A^{-}} + \mathrm{H^+} &&
K_\mathrm{a2} = \frac{[\mathrm{A^{-}}][\mathrm{H^+}]}{[\mathrm{A}]}
\end{aligned}\end{equation}
$$

中括号表示各个物质的摩尔浓度，通过变换可以将这些物质的摩尔浓度表示为 $\mathrm{[A]}$ 的倍数：

$$
\begin{equation}\begin{aligned}
\ &[\mathrm{A}] =\frac{K_\mathrm{a1}}{[\mathrm{H^+}]}[\mathrm{AH^+}]\\
&[\mathrm{A^{-}}]=\frac{K_\mathrm{a2}}{[\mathrm{H^+}]}[\mathrm{A}]=\frac{K_\mathrm{a2}K_\mathrm{a1}}{[\mathrm{H^+}]^2}[\mathrm{AH^+}]
\end{aligned}\end{equation}
$$

花色苷的总浓度为:

$$
\begin{equation}\begin{aligned}
\ [\mathrm{ACN}]_0&=[\mathrm{AH^+}]+[\mathrm{A}]+[\mathrm{A^{-}}]\\
&=\frac{[\mathrm{AH^+}]}{[\mathrm{H^+}]^2}([\mathrm{H^+}]^2+[\mathrm{H^+}]K_\mathrm{a1}+K_\mathrm{a1}K_\mathrm{a2})
\end{aligned}\end{equation}
$$

因此，各个物质的摩尔分数可以表示为：

$$
\begin{equation}\begin{aligned}
\chi_{\mathrm{AH^+}}&=\frac{[\mathrm{H^+}]}{D},\qquad
\chi_{\mathrm{A}}=\frac{[\mathrm{H^+}]K_\mathrm{a1}}{D},\qquad
\chi_{\mathrm{A^-}}=\frac{K_\mathrm{a1}K_\mathrm{a2}}{D}\\
D&=[\mathrm{H^+}]^2+[\mathrm{H^+}]K_\mathrm{a1}+K_\mathrm{a1}K_\mathrm{a2}
\end{aligned}\end{equation}
$$

一般来说，一定波长下的吸光度可以认为是体系中所有物质吸光度的线性叠加：

$$
\begin{equation}\begin{aligned}
A_{\mathrm{obs}}&=A_{\mathrm{AH^+}}+A_{\mathrm{A}}+A_{\mathrm{A^-}}\\
&=(\varepsilon_\mathrm{AH^+}\chi_\mathrm{AH^+}+\varepsilon_\mathrm{A}\chi_\mathrm{A}+\varepsilon_\mathrm{A^{-}}\chi_\mathrm{A^{-}})[\mathrm{ACN}]_0l\\
&=\frac{\varepsilon_\mathrm{AH^+}[\mathrm{H^+}]^2+\varepsilon_\mathrm{A}[\mathrm{H^+}]K_\mathrm{a1}+\varepsilon_\mathrm{A^{-}}K_\mathrm{a1}K_\mathrm{a2}}{[\mathrm{H^+}]^2+[\mathrm{H^+}]K_\mathrm{a1}+K_\mathrm{a1}K_\mathrm{a2}}[\mathrm{ACN}]_0l\\
&=\frac{a[\mathrm{H^+}]^2+b[\mathrm{H^+}]K_\mathrm{a1}+cK_\mathrm{a1}K_\mathrm{a2}}{[\mathrm{H^+}]^2+[\mathrm{H^+}]K_\mathrm{a1}+K_\mathrm{a1}K_\mathrm{a2}}\\
&=a\chi_\mathrm{A}+b\chi_\mathrm{A^-}+c\chi_\mathrm{A^{2-}}
\end{aligned}\end{equation}
$$

测定不同 pH 下的吸光度后可利用上述公式拟合，$a$、$b$ 、$c$、 $K_\mathrm{a1}$、$K_\mathrm{a2}$ 为拟合参数。

#### 2. Excel 中的拟合操作

将不同 pH 条件下测得的吸收光谱进行预处理（如筛选、取平均等），然后在多个波长处选取吸光度数据。

波长的选择通常应集中在那些随 pH 变化而吸光度变化显著的区域。

在excel按以下方式输入：

| pH   | \[H\]                 | $A_{实验值}$ | $A_{计算值}$ | $\chi_{\mathrm{A}}$ | $\chi_{\mathrm{A^-}}$ | $\chi_{\mathrm{A^{2-}}}$ |
| ---- | ------------------- | ------------ | ------------ | ------------------- | --------------------- | ------------------------ |
| 2    | $10^{-\mathrm{pH}}$ | 0.4          | 按公式输入   | 按公式输入          | 按公式输入            | 按公式输入               |
| 3    |                     | 0.6          |              |                     |                       |                          |

另外在旁边列出 $a$、$b$ 、$c$、 $K_\mathrm{a1}$、$K_\mathrm{a2}$ 等拟合参数，计算出所对应所有数据的$A_{实验值}$与$A_{计算值}$ 差的平方和（函数 SUMXMY2），拟合的目的是让这个值最小，就是说通过调整拟合参数使得计算值与实验值更接近。

需要打开excel的规划求解插件：文件-选项-加载项-管理（Excel加载项）-转到-勾选规划求解加载项-确定，然后可以在“数据”工作栏找到。

![solver](..\image\solver.png)

位置1 选$A_{实验值}$与$A_{计算值}$差的平方和；位置2 选最小值；位置3 选拟合参数

点击求解，即可获得优化后的位置3的参数

#### 3. 单组份光谱的获得（光谱分解）

这里不需要理解数学原理，能够做出来就可以，

已知

矩阵 ***D*** (m 行波长 × n 列样品)，数字为吸光度

矩阵 ***C*** (n 行样品 × l 列组分)，数字为摩尔分数

求解 矩阵 ***S*** (m 行波长 × l 列组分)，数字为吸光度

$$
\begin{equation}\begin{aligned}
D&=SC\\
S&=D[(C^TC)^-C^T]^T
\end{aligned}\end{equation}
$$

利用 excel 里面的 TRANSPOSE, MMULT, MINVERSE等函数可实现由 D 和 C 获得 S

可以带回去单组份的光谱乘以不同pH下的摩尔分数，与实验值进行对比，评估拟合结果。

#### 4. 光谱与颜色的转换

(1) 利用 [生成颜色表.xlsm](..\file\生成颜色表1.xlsm)

(2) 这里需要 Origin软件的 Chromaticity Diagram 插件，可以在Origin自带的商店里安装。

将吸收光谱换算为透射光谱 $T=10^{-A}$ 拷入Origin，选中数据，点击插件，点击第二个按钮PL，选择Reflectance or Transmittance (0-1), D65, CIE 1964 10°，其余随便

在获得的图里右键点击数据点，跳转到Book，里面有计算所得 L\*, a\*, b\* 值，可以通过多种渠道将 L\*, a\*, b\* 转化为 RGB，这里介绍 Excel 里代码转化的方法。

excel 左下角Sheet名称处右键-查看代码，左上角插入模块，把下面代码拷入，获得 LABtoRGB 函数，输入 LABtoRGB(L\*, a\*, b\*) 可得 (R, G, B) 在一个单元格里。注意LABtoRGB可引用单元格数据，单只能逐个引用，比如 LABtoRGB(A1, A2, A3)，不能 LABtoRGB(A1:A3)。

```
Function LABtoRGB(L As Double, a As Double, labB As Double) As String
    Dim X As Double, Y As Double, Z As Double
    Dim R As Double, G As Double, B As Double
    Dim var_X As Double, var_Y As Double, var_Z As Double

    ' 参考白点 D65
    Dim ref_X As Double: ref_X = 95.047
    Dim ref_Y As Double: ref_Y = 100#
    Dim ref_Z As Double: ref_Z = 108.883

    ' LAB -> XYZ
    var_Y = (L + 16) / 116
    var_X = a / 500 + var_Y
    var_Z = var_Y - labB / 200

    If var_Y ^ 3 > 0.008856 Then
        Y = var_Y ^ 3
    Else
        Y = (var_Y - 16 / 116) / 7.787
    End If

    If var_X ^ 3 > 0.008856 Then
        X = var_X ^ 3
    Else
        X = (var_X - 16 / 116) / 7.787
    End If

    If var_Z ^ 3 > 0.008856 Then
        Z = var_Z ^ 3
    Else
        Z = (var_Z - 16 / 116) / 7.787
    End If

    X = X * ref_X
    Y = Y * ref_Y
    Z = Z * ref_Z

    ' XYZ -> RGB (sRGB)
    R = X * 0.032406 + Y * -0.015372 + Z * -0.004986
    G = X * -0.009689 + Y * 0.018758 + Z * 0.000415
    B = X * 0.000557 + Y * -0.00204 + Z * 0.01057

    ' Gamma校正
    R = IIf(R > 0.0031308, 1.055 * (R ^ (1 / 2.4)) - 0.055, 12.92 * R)
    G = IIf(G > 0.0031308, 1.055 * (G ^ (1 / 2.4)) - 0.055, 12.92 * G)
    B = IIf(B > 0.0031308, 1.055 * (B ^ (1 / 2.4)) - 0.055, 12.92 * B)

    ' 限制在0-1之间
    R = WorksheetFunction.Min(WorksheetFunction.Max(R, 0), 1)
    G = WorksheetFunction.Min(WorksheetFunction.Max(G, 0), 1)
    B = WorksheetFunction.Min(WorksheetFunction.Max(B, 0), 1)

    ' 转换成0-255
    R = R * 255
    G = G * 255
    B = B * 255

    ' 返回字符串 "R,G,B"
    LABtoRGB = Round(R) & "," & Round(G) & "," & Round(B)
End Function
```

选中有 (R, G, B) 的单元格，在 excel 左下角Sheet名称处右键-查看代码，拷入下面的代码后运行，可以将单元格及其中文字转化为对应颜色。

```
Sub SetColor_SelectedCells()
    Dim cell As Range
    Dim R As Integer, G As Integer, B As Integer
    Dim arr As Variant

    ' 遍历选中区域内的每个单元格
    For Each cell In Selection
        If InStr(cell.Value, ",") > 0 Then
            arr = Split(cell.Value, ",")
            If UBound(arr) = 2 Then
                R = CInt(arr(0))
                G = CInt(arr(1))
                B = CInt(arr(2))
                
                ' 设置背景色
                cell.Interior.Color = RGB(R, G, B)
                
                ' 设置字体颜色
                cell.Font.Color = RGB(R, G, B)
            End If
        End If
    Next cell
End Sub
```



