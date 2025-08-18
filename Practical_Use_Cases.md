## 実践的なスクリプト

[気象庁のウェブページ](https://www.data.jma.go.jp/stats/etrn/index.php)から[北見観測所](https://www.data.jma.go.jp/stats/etrn/view/daily_a1.php?prec_no=17&block_no=0074&year=2025&month=7)のデータをスクレーピングする。

### 2025 年 7 月のデータを表示する URL

<p style="margin-left:20px">
https://www.data.jma.go.jp/stats/etrn/view/daily_a1.php<br>
<span style="color:red">?</span>prec_no=17<br>
<span style="color:red">&</span>block_no=0074<br>
<span style="color:red">&</span>year=2025<br>
<span style="color:red">&</span>month=7
</p>

&emsp;&emsp;※ のぼっこ [気象庁の都府県・地方と地点の各コードを書き出してみた話](https://zenn.dev/nobokko/articles/idea_jma_prec_no_and_block_no#%E5%9C%B0%E7%82%B9%E3%81%AE%E3%82%B3%E3%83%BC%E3%83%89%E8%A1%A8)

### HTML（抜粋）

```html
  1	
  2	<!doctype html>
  3	<html lang="ja">
  4	<head>
            :
 19	</head>
 20	<body>
            :
 77	<table id="tablefix1" class="data2_s">
 78	<tr class="mtx">
            :   1 行目（項目名）
 86	</tr>
 87	<tr class="mtx">
            :   2 行目（同）
 93	</tr>
 94	<tr class="mtx">
            :   3 行目（同）
 97	</tr>
 98	<tr class="mtx" style="text-align:right;"><td style="white-space:nowrap"><div clas
s="a_print"><a href="hourly_a1.php?prec_no=17&block_no=0074&year=2025&month=07&day=1&view=
p1">1</a></div></td><td class=data_0_0>2.0</td><td class=data_0_0>2.0</td><td class=data_0
_0>1.5</td><td class=data_0_0>23.8</td><td class=data_0_0>30.6</td><td class=data_0_0>19.9
</td><td class=data_0_0>81</td><td class=data_0_0>54</td><td class=data_0_0>2.3</td><td cl
ass=data_0_0>5.3</td><td class=data_0_0 style="text-align:center">北東</td><td class=dat
a_0_0>8.5</td><td class=data_0_0 style="text-align:center">北東</td><td class=data_0_0 s
tyle="text-align:center">南西</td><td class=data_0_0>4.9</td><td class=data_0_0>0</td><t
d class=data_0_0>0</td></tr>
    :   5 行目以降
129	</table>
    :
150	....... </body>
151	</html>
```

### VBA

```vb
 1	' 気象庁のウェブページから北見観測所のデータをスクレーピングする
 2	
 3	Const URL_DAILY = "https://www.data.jma.go.jp/stats/etrn/view/daily_a1.php?"
 4	Const INVALID = -999
 5	'
 6	
 7	Private Sub CommandButton1_Click()
 8	
 9	    Dim Caps As Object, row As Object, col As Object
10	
11	    fname = "kitami_daily.csv"
12	    URL_PREFIX = URL_DAILY & "prec_no=17&block_no=0074" '北見
13	
14	    dt_init = #1/1/2023# '取得開始日
15	    dt_exit = #1/1/2025# '取得終了の翌日
16	
17	    Open ThisWorkbook.Path & "\" & fname For Output As #1
18	
19	    Print #1, "date,rain_24h,rain_60m,rain_10m,temp_ave,temp_max,temp_min"
20	
21	    With New SeleniumVBA.WebDriver
22	
23	        .StartChrome
24	        Set Caps = .CreateCapabilities
25	        Caps.RunInvisible
26	
27	        .OpenBrowser Caps
28	
29	        days = dt_exit - dt_init
30	        dt = dt_init
31	        Do
32	            Application.StatusBar = "Progress : " & _
33	                                    Format(100 * (dt - dt_init) / days, "0.0") & " %"
34	
35	            .NavigateTo URL_PREFIX & "&year=" & Year(dt) & "&month=" & Month(dt)
36	            .Wait 200
37	
38	            With .FindElementByID("tablefix1")
39	
40	                row_no = 1
41	                For Each row In .FindElementsByTagName("tr")
42	
43	                    If row_no > 3 Then
44	                        
45	                        col_no = 1
46	                        For Each col In row.FindElementsByTagName("td")
47	
48	                            If col_no = 1 Then
49	
50	                                buf = Format(dt, "yyyy-mm-dd") 'ISO-8601
51	                                
52	                            Else
53	
54	                                data = col.GetInnerHTML
55	
56	                                If IsNumeric(data) Then
57	                                    buf = buf & Format(Val(data), ",0.0")
58	                                Else
59	                                    buf = buf & Format(INVALID, ",#")
60	                                End If
61	
62	                                If col_no = 7 Then Exit For
63	                            End If
64	
65	                            col_no = col_no + 1
66	                        Next col
67	
68	                        Print #1, buf
69	
70	                        dt = dt + 1
71	                        If dt = dt_exit Then Exit Do
72	                    End If
73	
74	                    row_no = row_no + 1
75	                Next row
76	
77	            End With
78	        Loop
79	
80	        .CloseBrowser
81	        .Shutdown
82	
83	    End With
84	
85	    Close #1
86	    Set Caps = Nothing
87	
88	    Application.StatusBar = ""
89	
90	End Sub
```

### kitami_daily.csv

<p style="margin-left:20px">
date,rain_24h,rain_60m,rain_10m,temp_ave,temp_max,temp_min<br>
2023-01-01,0.0,0.0,0.0,-11.1,-1.9,-18.2<br>
2023-01-02,0.0,0.0,0.0,-13.3,-7.5,-20.9<br>
2023-01-03,0.0,0.0,0.0,-17.7,-9.5,-24.5<br>
2023-01-04,0.0,0.0,0.0,-16.3,-7.6,-22.2<br>
&emsp;&emsp;&emsp;:<br>
2024-12-27,0.0,0.0,0.0,-7.4,-4.0,-14.4<br>
2024-12-28,0.0,0.0,0.0,-5.6,-1.7,-16.0<br>
2024-12-29,0.0,0.0,0.0,-8.0,-1.2,-13.5<br>
2024-12-30,0.0,0.0,0.0,-10.3,-5.2,-16.8<br>
2024-12-31,4.5,1.0,0.5,-9.7,-4.3,-18.4
</p>

----
