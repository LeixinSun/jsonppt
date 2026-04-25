
## 一、PPT 文件本身是什么

`.pptx` 文件其实是一个 **zip 压缩包**,里面装的是一堆 XML 文件。你把 output.pptx 改成 output.zip 解压就能看到。结构大致是:

```
output.pptx (其实是 zip)
├── ppt/
│   ├── slides/slide1.xml      ← 第一页的所有内容
│   ├── slides/slide2.xml      ← 第二页
│   ├── media/image1.jpg       ← 嵌入的图片
│   └── theme/theme1.xml       ← 配色、字体
└── [Content_Types].xml
```

每一页的 XML 长这样(简化版):

```xml
<p:sp>  <!-- 一个形状 -->
  <p:spPr>
    <a:xfrm>
      <a:off x="914400" y="914400"/>      <!-- 位置 -->
      <a:ext cx="2743200" cy="1828800"/>  <!-- 尺寸 -->
    </a:xfrm>
    <a:prstGeom prst="roundRect"/>        <!-- 圆角矩形 -->
    <a:solidFill><a:srgbClr val="6D2E46"/></a:solidFill>
  </p:spPr>
  <p:txBody>
    <a:p><a:r><a:t>01 · 产地</a:t></a:r></a:p>  <!-- 文字内容 -->
  </p:txBody>
</p:sp>
```

也就是说 —— **PPT 里没有"魔法",每个方框、文字、图片都只是 XML 里的几行坐标和样式定义**。坐标用 EMU 单位(914400 EMU = 1 英寸)。

## 二、JSON → PPT 是怎么翻译的

我没有让 AI "凭感觉画图",而是做了一个机械的翻译:

```
你的 JSON          →    Python 代码           →    XML 节点         →    .pptx
{                       slide.shapes               <p:sp>
  "type":"flow_node",   .add_shape(                  <p:spPr>
  "x": 5.0,              ROUNDED_RECTANGLE,           <a:off x="...">
  "y": 2.1,              Inches(5.0),                 ...
  ...                    Inches(2.1),
}                        ...
                       )
```

**关键工具是 `python-pptx`** —— 这是一个 Python 库,把"加形状/加文字/加图片"这种命令翻译成正确的 XML,再打包成 zip。它不涉及任何 AI,纯粹是机械的文件构造。

`render.py` 做的事就是:

1. 读 JSON,拿到一个元素列表
2. 对每个元素,看 `type` 字段,分发给对应的处理函数(`add_text`/`add_image`/`add_flow_node` ...)
3. 处理函数调用 python-pptx 的 API,在某个坐标加一个形状/文字/图片
4. 全部加完,保存成 .pptx

```
spec.json  ──读取──▶  render.py  ──python-pptx──▶  output.pptx
                       │
                       ├─ 元素 1: 矩形       → add_rect()
                       ├─ 元素 2: 文字       → add_text()
                       ├─ 元素 3: 图片       → add_image()
                       ├─ 元素 4: 流程节点   → add_flow_node()
                       ├─ 元素 5: 箭头       → add_flow_arrow()
                       └─ ...
```

## 三、AI(我)在这里到底做了什么

这一点很关键 —— **生成最终 PPT 的不是 AI,是确定性的代码**。AI 的作用其实只在两个环节:

| 环节 | 谁做的 | 是否确定性 |
|---|---|---|
| 想主题、想配色、设计版式 | AI(我) | 创造性,每次可能不同 |
| 写出 JSON(布局蓝图) | AI(我) | 创造性 |
| 写 render.py(翻译器) | AI(我),但只写一次 | 之后是死的代码 |
| **JSON → .pptx 文件** | **render.py(纯代码)** | **完全确定性,同样 JSON 永远产出同样 PPT** |

这种架构的好处:
- **可复现**:同一份 JSON + 图片,跑一万次结果完全一样
- **可控**:不会出现 AI "幻觉"出一个奇怪的版式
- **易修改**:你看到 PPT 哪里不对,直接改 JSON 里的数字,不用重新让 AI 生成
- **可批量**:有 100 页内容?写个脚本生成 100 份 JSON,一次跑完

## 类比

可以这样想:

> JSON 是建筑蓝图,render.py 是施工队,python-pptx 是工具箱,.pptx 是建好的房子。
> 
> AI 的角色是 **建筑师** —— 听你需求、画蓝图。
> 真正盖房子的还是施工队,而且每次按同一份蓝图盖出来的房子完全一样。

如果不用 JSON 这一层,直接让 AI 生成 .pptx 的二进制内容,那才是真"AI 合成"——但既不可控也不可复现,也是为什么市面上 AI PPT 工具大都内部走类似的"AI 出结构化数据 → 模板渲染"路线。
