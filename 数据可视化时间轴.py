from pyecharts import options as opts
from pyecharts.charts import Bar, Timeline
from pyecharts.commons.utils import JsCode
from pyecharts.faker import Faker
import random
a=['1月','2月','3月','4月','5月','6月''7月','8月','9月','10月','11月','12月']
# a=['1月','2月','3月']
# b1=[0,0,0,0,0,0,0,0,0,0,0,0,262,290,883,1160,96,680,155,1841,555,407,2271,1292,280,210,1675,1421,0,0,0,0,0,0,0,0]
# b1=[1, 5, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]

# b2=[0,0,201,682,919,569,966,576,91,81,158,46,293,221,209,652,393,484,395,524,658,783,1168,1609,1760,1934,2018,1150,0,0,0,0,0,0,0,0]
# b2=[5,6,7,8,9,10,11,12,13,14,15,16]
# print(len(b1))
datas={}
x = Faker.choose()
tl = Timeline()
for i in range(2019, 2022):
    bar = (
        Bar()
        .add_xaxis(a)
        .add_yaxis("病原",[random.randint(200, 2000) for _ in range(12)])
        .add_yaxis("抗体",[random.randint(200, 2000) for _ in range(12)])
        .set_global_opts(
            title_opts=opts.TitleOpts("病原{}年数量 ".format(i)),
            graphic_opts=[
                opts.GraphicGroup(
                    graphic_item=opts.GraphicItem(
                        rotation=JsCode("Math.PI / 4"),
                        bounding="raw",
                        right=100,
                        bottom=110,
                        z=100,
                    ),
                    children=[
                        opts.GraphicRect(
                            graphic_item=opts.GraphicItem(
                                left="center", top="center", z=100
                            ),
                            graphic_shape_opts=opts.GraphicShapeOpts(
                                width=400, height=50
                            ),
                            graphic_basicstyle_opts=opts.GraphicBasicStyleOpts(
                                fill="rgba(0,0,0,0.3)"
                            ),
                        ),
                        opts.GraphicText(
                            graphic_item=opts.GraphicItem(
                                left="center", top="center", z=100
                            ),
                            graphic_textstyle_opts=opts.GraphicTextStyleOpts(
                                text="实验室{}年检测数量".format(i),
                                font="bold 26px Microsoft YaHei",
                                graphic_basicstyle_opts=opts.GraphicBasicStyleOpts(
                                    fill="#fff"
                                ),
                            ),
                        ),
                    ],
                )
            ],
        )
    )
    tl.add(bar, "{}年".format(i))
tl.render(r'D:\Users\Administrator\Desktop\1.html')