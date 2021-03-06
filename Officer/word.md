Word
================
Lu Yiyun
2020/1/1

  - [1. Quick start](#quick-start)
  - [2. 增加元素](#增加元素)
  - [3. 光标操作](#光标操作)
      - [4.1 替换内容](#替换内容)
      - [4.2 搜索替换](#搜索替换)
  - [5. 对节（sections）的操作](#对节sections的操作)

# 1\. Quick start

1.  创建一个rdocx对象，通过`read_docx(path =
    NULL)`，参数`path`如果为空，则代表创建了一个空的docx文件，如果指定了一个已存在的docx文件，则创建的rdocx对象中会保存这个docx文件的内容，便于进一步修改；

<!-- end list -->

``` r
library(officer)

my_doc <- read_docx()
str(my_doc)
## List of 2
##  $ package_dir   : chr "C:\\Users\\hasee\\AppData\\Local\\Temp\\Rtmp8Szk9C\\file444868c216d5"
##  $ doc_properties:List of 2
##   ..$ data: chr [1:10, 1:4] "dc" "dc" "dc" "cp" ...
##   .. ..- attr(*, "dimnames")=List of 2
##   .. .. ..$ : chr [1:10] "title" "subject" "creator" "keywords" ...
##   .. .. ..$ : chr [1:4] "ns" "name" "attrs" "value"
##   ..$ ns  : Named chr [1:5] "http://schemas.openxmlformats.org/package/2006/metadata/core-properties" "http://purl.org/dc/elements/1.1/" "http://purl.org/dc/dcmitype/" "http://purl.org/dc/terms/" ...
##   .. ..- attr(*, "names")= chr [1:5] "cp" "dc" "dcmitype" "dcterms" ...
##   ..- attr(*, "class")= chr "core_properties"
##  - attr(*, "class")= chr "rdocx"
```

我们会进一步对其进行修改，比如增加一些段落等等。这里提前告知一下，我们可以通过函数`style_info`作用于任何一个rdocx对象来看到我们可以操作的内容，以及这些内容的可选格式：

``` r
styles_info(my_doc)
##    style_type            style_id             style_name is_custom
## 1   paragraph              Normal                 Normal     FALSE
## 2   paragraph              Titre1              heading 1     FALSE
## 3   paragraph              Titre2              heading 2     FALSE
## 4   paragraph              Titre3              heading 3     FALSE
## 5   character      Policepardfaut Default Paragraph Font     FALSE
## 6       table       TableauNormal           Normal Table     FALSE
## 7   numbering         Aucuneliste                No List     FALSE
## 8   character              strong                 strong      TRUE
## 9   paragraph            centered               centered      TRUE
## 10      table       tabletemplate         table_template      TRUE
## 11      table Listeclaire-Accent2    Light List Accent 2     FALSE
## 12  character           Titre1Car            Titre 1 Car      TRUE
## 13  character           Titre2Car            Titre 2 Car      TRUE
## 14  character           Titre3Car            Titre 3 Car      TRUE
## 15  paragraph        graphictitle          graphic title      TRUE
## 16  paragraph          tabletitle            table title      TRUE
## 17      table       Professionnel     Table Professional     FALSE
## 18  paragraph                 TM1                  toc 1     FALSE
## 19  paragraph                 TM2                  toc 2     FALSE
## 20  paragraph       Textedebulles           Balloon Text     FALSE
## 21  character    TextedebullesCar    Texte de bulles Car      TRUE
## 22  character         referenceid           reference_id      TRUE
##    is_default
## 1        TRUE
## 2       FALSE
## 3       FALSE
## 4       FALSE
## 5        TRUE
## 6        TRUE
## 7        TRUE
## 8       FALSE
## 9       FALSE
## 10      FALSE
## 11      FALSE
## 12      FALSE
## 13      FALSE
## 14      FALSE
## 15      FALSE
## 16      FALSE
## 17      FALSE
## 18      FALSE
## 19      FALSE
## 20      FALSE
## 21      FALSE
## 22      FALSE
```

2.  往rdocx中创建内容。这里有多种函数来为其添加不同的内容，其中最常用的是`body_add_par`，其中`par`表示的是paragraph（段落）。也可以添加图片：

<!-- end list -->

``` r
# 创建一个临时的图片
src <- tempfile(fileext = ".png")
png(filename = src, width = 5, height = 6, units = 'in', res = 300)
barplot(1:10, col = 1:10)
dev.off()
## png 
##   2
# 将这个图片和一些文字加入到rdocx中
my_doc <- my_doc %>% 
  body_add_img(src = src, width = 5, height = 6, style = "centered") %>% 
  body_add_par("Hello world!", style = "Normal") %>% 
  body_add_par("", style = "Normal") %>% # blank paragraph
  body_add_table(iris, style = "table_template")
```

默认这些内容都添加在文件的后面，至于如何在任意位置添加内容，就需要设计cursor的操作，这在下面会讲到。

3.  最后使用`print(path=..)`来将rdocx对象变成真实的docx文件：

<!-- end list -->

``` r
print(my_doc, "./first_example.docx")
```

最后的结果展示![office2](./images/officer2.png)

# 2\. 增加元素

这里有两类函数来增加元素：

  - **增加的内容是top container，使用的是`body_add_*`函数**：
      - `body_add_par(x, value, style = NULL, pos =
        "after")`，加入一个文字段落，`pos`表示的是加入的内容在**光标（cursor）选择的内容**之前还是之后（一般选择的内容就是上一次创建的top
        container，这里的光标类似vim，不是在两个字符之间，而是选择（覆盖）了一个对象，而且和vim不同的是，其选择的一般是一个段落），
        `style`是段落的风格名称，是`styles_info`输出的“style\_name”一列；
      - `body_add_img(x, src, style = NULL, width, height, pos =
        "after")`，`src`是file的路径；
      - `body_add_table(x, value, style = NULL, pos = "after", header =
        TRUE, first_row = TRUE, first_column = FALSE, last_row = FALSE,
        last_column = FALSE, no_hband = FALSE, no_vband = TRUE)`，添加表格；
      - `body_add_break(x, pos = "after")`，添加一个分页符；
      - `body_add_toc(x, level = 3, pos = "after", style = NULL,
        separator = ";")`，添加一个目录；
      - `body_add_gg(x, value, width = 6, height = 5, style = NULL,
        ...)`，添加一个ggplot对象到word中，其作为一个png图像被插入到其中，`...`表示会输入到`png()`中的参数；
      - `body_add_fpar(x, value, style = NULL, pos =
        "after")`，增加fpar到word中；
      - `body_add_blocks(x, blocks, pos =
        "after")`，将`block_list`创建的`block`对象（表示的是多个段落或图片的合集）添加到rdocx中；
      - `flextable::body_add_flextable(x, value, align = "center", pos =
        "after", split =
        FALSE)`，这来自于`flextable`，使得我们可以添加更加漂亮的表格到word文档中；
  - **将文字或图片插入到已经存在的段落中，使用的是`slip_in_*`系列函数**，其只能做到在存在的段落的前面或后面添加内容，添加的内容还是算作在这个段落中的，可能实现在一个段落中出现不同格式的文字（比如实现表头和图片标题的时候是非常需要的）：
      - `slip_in_img(x, src, style = NULL, width, height, pos =
        "after")`，添加图片，这里的`pos`表示的意思和上面的一致，只是增加的内容不会出现新的段落，而是和当前选中的段落组合在一起；
      - `slip_in_text(x, str, style = NULL, pos = "after", hyperlink =
        NULL)`，添加文字，这里可以将文字变成超链接；
      - `slip_in_seqfield(x, str, style = NULL, pos = "after")`，添加seq
        field到段落中；
      - `slip_in_footnote(x, style = NULL, blocks, pos =
        "after")`，添加脚注，其中脚注的内容是blocks定义的；
      - `slip_in_column_break(x, pos = "before")`，添加分栏符；
      - `slip_in_xml(x, str, pos)`，添加一个wml string到rdocx中；

> 一般来说，是一个`body_add_*`后面跟几个`slip_in_*`函数，`slip_in_*`负责对创建的top
> container进行修改，在头和尾加入一些特殊格式的内容，其用法可以见下面`slip_in_seqfield`的示例。

这是关于`body_add_*`的演示，其结果在[Introduction](./introduction.md)中有展示，就不再展示了：

``` r
library(ggplot2)
library(flextable)

# body_add_*应用试验
gg <- ggplot(data = iris, aes(Sepal.Length, Petal.Length)) + 
  geom_point()
ft <- qflextable(head(iris))
read_docx() %>% 
  body_add_par(value = "Table of content", style = "heading 1") %>% 
  body_add_toc(level = 2) %>% 
  body_add_break() %>% 

  body_add_par(value = "dataset iris", style = "heading 2") %>% 
  body_add_flextable(value = ft ) %>% 
  
  body_add_par(value = "plot examples", style = "heading 1") %>% 
  body_add_gg(value = gg, style = "centered" ) %>% 

  print(target = "./body_add_demo.docx")
```

这是关于`slip_in_*`函数的展示，其结果展示在下面：

``` r
# slip_in_*应用试验
img.file <- file.path( R.home("doc"), "html", "logo.jpg" )
read_docx() %>%
  body_add_par("R logo: ", style = "Normal") %>%
  slip_in_img(src = img.file, style = "strong", 
              width = .3, height = .3, pos = "after") %>% 
  slip_in_text(" - This is ", style = "strong", pos = "before") %>% 
  slip_in_column_break(pos = "after") %>% 
  slip_in_seqfield(str = "SEQ Figure \u005C* ARABIC",
    style = 'strong', pos = "before") %>% 
  print(target = "./slip_in_demo.docx")
```

![officer3](./images/officer3.png)

以下是关于`slip_in_seqfield`的示例，但也表达了如何使用这些函数：

``` r
x <- read_docx() %>%
  body_add_par("Time is: ", style = "Normal") %>%
  slip_in_seqfield(
    str = "TIME \u005C@ \"HH:mm:ss\" \u005C* MERGEFORMAT",
    style = 'strong') %>%

  body_add_par(" - This is a figure title", style = "centered") %>%
  slip_in_seqfield(str = "SEQ Figure \u005C* roman",
    style = 'Default Paragraph Font', pos = "before") %>%
  slip_in_text("Figure: ", style = "strong", pos = "before") %>%

  body_add_par(" - This is another figure title", style = "centered") %>%
  slip_in_seqfield(str = "SEQ Figure \u005C* roman",
    style = 'strong', pos = "before")  %>%
  slip_in_text("Figure: ", style = "strong", pos = "before") %>%
  body_add_par("This is a symbol: ", style = "Normal") %>%
  slip_in_seqfield(str = "SYMBOL 100 \u005Cf Wingdings",
    style = 'strong')
```

![officer4](./images/officer4.png)

# 3\. 光标操作

光标（cursor）是可以操作的，我们可以通过对光标的操作来改变我们现在正选择的内容。上面的函数我们也看到了普遍有参数`pos`，这个参数表示的是在当前选择的内容的前面还是后面添加这次新的内容，其有以下3个选项：

  - `before`;
  - `after`;
  - `on`，将会替换选择的内容。

可以操作光标的函数有：

  - `cursor_begin(x)`，将光标放在文档的第一个对象（一般指的就是第一个段落或表格）上；
  - `cursor_bookmark(x, id)`，将光标放置在设定好的书签的对象上；
  - `cursor_end(x)`，将光标放置在文档的最后一个对象上；
  - `cursor_reach(x,
    keyword)`，将光标放置在包含`keyword`的第一个对象上，这个`keyword`可以是正则表达式；
  - `cursor_forward(x)`，将光标前移一个对象；
  - `cursor_backward(x)`，将光标后移一个对象。

这些函数主要用来移动光标选中特定的段落，然后我们再使用另外的函数，就可以修改光标选中的段落了。

为了便于演示光标的操作，我们先创建一个文档：

``` r
read_docx() %>%
  body_add_par("paragraph 1", style = "Normal") %>%
  body_add_par("paragraph 2", style = "Normal") %>%
  body_add_par("paragraph 3", style = "Normal") %>%
  body_add_par("paragraph 4", style = "Normal") %>%
  body_add_par("paragraph 5", style = "Normal") %>%
  body_add_par("paragraph 6", style = "Normal") %>%
  body_add_par("paragraph 7", style = "Normal") %>%
  print(target = "./init_doc.docx" )
```

它的样式是：

![officer5](./images/officer5.png)

以下是通过光标进行一系列操作：

``` r
doc <- read_docx(path = "./init_doc.docx") %>%

  # 移除第一个段落（默认模板，第一个段落是一个空段落，移除它）
  cursor_begin() %>% body_remove() %>%

  # 移动到包含paragraph 4的段落，然后在其前面加上字符This is
  cursor_reach(keyword = "paragraph 4") %>%
  slip_in_text("This is ", pos = "before", style = "Default Paragraph Font") %>%

  # 再前移一个对象（paragraph 5），在其后加入一个新的段落
  cursor_forward() %>%
  body_add_par("The section stop here", style = "Normal") %>%
  # body_end_section(landscape = TRUE, continuous = FALSE) %>%

  # 在文档末尾加入一个新的段落，指示文档结束
  cursor_end() %>%
  body_add_par("The document ends now", style = "Normal")

print(doc, "./cursor.docx")
```

最后的结果是： ![officer6](./images/officer6.png) \# 4. 内容操作

以上的实验中包含了一个操作内容的函数`body_remove`，即删除光标选中的段落。实际上`officer`提供了一系列进行内容操作的函数，现在列举如下：

  - `body_remove(x)`;
  - `body_bookmark(x, id)`，为某个段落打上书签；
  - 还有一系列进行内容替换的函数，后面将进行详细说明。

## 4.1 替换内容

替换掉整个段落，就是前面说的，使用`cursor_*`来确定段落位置，使用`body_add_*`中的`pos="on"`来替换。

``` r
str1 <- "Lorem ipsum dolor sit amet, consectetur adipiscing elit. " %>% 
  rep(20) %>% paste(collapse = "")
str2 <- "Drop that text" 
str3 <- "Aenean venenatis varius elit et fermentum vivamus vehicula. " %>% 
  rep(20) %>% paste(collapse = "")

read_docx()  %>% 
  body_add_par(value = str1, style = "Normal") %>% 
  body_add_par(value = str2, style = "centered") %>% 
  body_add_par(value = str3, style = "Normal") %>% 
  print("./replace_template.docx")
```

替换之前： ![officer7](./images/officer7.png)

``` r
read_docx("./replace_template.docx") %>% 
  cursor_reach("that text") %>% 
  body_add_par("This is a new paragraph.", style = "centered", pos = "on") %>% 
  print("./replace_doc.docx")
```

替换之后： ![officer8](./images/officer8.png)

## 4.2 搜索替换

除了上面通过`cursor_*`和`body_add_*`两者组合进行替换的方式外，还有一种是直接使用一个函数来完成搜索替换的功能，这主要有一组函数来完成：

  - `body_replace_text_at_bkm(x, bookmark, value)`
  - `body_replace_img_at_bkm(x, bookmark,
    value)`，这两个函数可以将标有书签`bookmark`的段落替换为`value`；
  - `body_replace_all_text(x, old_value, new_value, only_at_cursor =
    FALSE, warn = TRUE, ...)`，这个函数可以做到在当前段落（`only_at_cursor=TRUE`）或全文
    （`only_at_cursor=FALSE`）内搜索`old_value`替换成`new_value`；

以上三个函数还有其`headers`、`footers`版本。

``` r
# 先构建一个要进行修改的文件
my_doc <- read_docx() %>% 
  body_add_par("paragraph one") %>% 
  body_bookmark("one") %>% 
  
  body_add_par("paragraph two") %>% 
  body_bookmark("two") %>% 
  
  body_add_par("paragraph three")
  
print(my_doc, "./search_replace.docx")
```

![officer9](./images/officer9.png)

``` r
my_doc %>% 
  # 替换bookmark == two
  body_replace_text_at_bkm("two", "replaced two") %>% 
  
  # 替换第三段的three
  body_replace_all_text("three", "THREE", only_at_cursor = TRUE) %>% 
  
  # 替换所有的paragraph
  body_replace_all_text("paragraph", "PARAgraph", only_at_cursor = FALSE) %>% 

  print("./search_replaced.docx")
```

![officer10](./images/officer10.png)

# 5\. 对节（sections）的操作

节是用来隔离两个部分的页面格式的，可以在同一页，也可以多页属于同一节。在word中分页符只有一种，而分节符有多种。
