# export-excel-in-one-line
前端excel导出工具

# 使用方式
1、安装依赖

```javascript
npm install xlsx-js-style
```
2、复制代码文件exportExcel.js至工程目录

[https://github.com/EnthuDai/export-excel-in-one-line](https://github.com/EnthuDai/export-excel-in-one-line)

3、在引入excel.js后调用

```javascript
Excel.export(columns, dataSource, '导出文件名')
```

4、代码demo

![代码示例](https://img-blog.csdnimg.cn/d6e8011845724f6583a84d339509de23.jpeg#pic_center)


5、效果
| 页面 | excel |
|--|--|
| ![在这里插入图片描述](https://img-blog.csdnimg.cn/691f19bfe3214afa97ac64669abd91b2.png#pic_center =500x)| ![在这里插入图片描述](https://img-blog.csdnimg.cn/b8fad81a8d9e43d9bfe1a930991b4803.png#pic_center)|
# 适用范围
对于使用vue ant-design 组件库中table组件的场景，可直接将table的 __columns__ 和 __data-source__、 __导出文件名称__ 作为参数传入export方法，调用即可导出相同格式的excel文件。
```javascript
Excel.export(this.demoColumn, this.demoData, '测试数据')
```

对于其他场景，需要对数据格式适配至ant-design table相同格式，具体为：

 1. 表头格式需符合以下条件
  - 标题的key为 *title*
  - 数据字段key为 *dataIndex*
  - 子表头key为 *children*
  
   如：

```javascript
	[
        {title:'类别',dataIndex:'type'},
        {title:'总计',children:[
            {title:'销量',children:[
                {title:'今天', dataIndex:'today'},
                {title:'昨天', dataIndex:'yesterday'}]
            }
          ]
        },
        {title:'趋势',children:[
            {title:'上涨率',dataIndex:'raise'},
            {title:'上涨金额', dataIndex:'raiseAmount'}
          ]
        }
      ]
```

 2. 数据格式格式需符合以下条件
  - 数据项key为表头格式中对应项 *dataIndex* 的值
  - 子数据key为*children* 
  
   如：
   

```javascript
	[
        {type:'笔', today:100, yesterday: 200, raise:'20%', raiseAmount:20, children:[
            {type:'毛笔',today:50, yesterday: 100, raise:'20%', raiseAmount:10},
            {type:'钢笔',today:50, yesterday: 100, raise:'20%', raiseAmount:10}
          ]},
        {type:'墨', today:100, yesterday: 200, raise:'20%', raiseAmount:20},
        {type:'纸', today:100, yesterday: 200, raise:'20%', raiseAmount:20},
        {type:'砚', today:100, yesterday: 200, raise:'20%', raiseAmount:20},
      ]
```

# 实现原理
原理基本参考了[使用xlsx.js导出有复杂表头的excel](https://blog.csdn.net/seeflyliu/article/details/109476804)这篇文章，其该文合并表头方法*doMerges* 存在bug，实测中会出现问题。所以该组件中使用了树中递归处理的算法计算合并项，解决问题的同时也提高了代码的简洁程度。
 实现过程：
 1. 根据表头描述 *columns* 生成全为空的表头二维数组，二维数组行数为 *columns* 中子项树的最深深度，列数为 *columns* 中所有子项树的叶子节点数之和。分别通过 columnHeight(*columns*)、columnWidth(*columns*)方法递归求得。
 ![在这里插入图片描述](https://img-blog.csdnimg.cn/b4ed53790e254680978517acbd4cfd38.png#pic_left)
 2. 将 *columns* 中title填入对应位置，也是循环+递归实现。此间分两种情况
 	1. 无children的叶子节点 

		```javascript
		{title:'类别',dataIndex:'type'}
		```
		在数组左上角第一项填入 *title*，合并单元格时需要向下合并所有单元格，记录下合并的起始和终点项的偏移量 **{s:{r:0,c:0},e:{r:0,c:2}}**
		
		![在这里插入图片描述](https://img-blog.csdnimg.cn/c3203936c8364250944bf8a6b9ef20a9.png#pic_left)
		2.有children的节点
		

		```javascript
		 {
		     title:'总计',children:[
               {title:'销量',children:[
                   {title:'今天', dataIndex:'today'},
                   {title:'昨天', dataIndex:'yesterday'}]
            }
          ]
        }
		```
		![在这里插入图片描述](https://img-blog.csdnimg.cn/759c17fcc26d480a976d2d87b33109ae.png#pic_left)
		在二维数组剩余的部分（红框区域）中，左上角第一项填入title，并记录下横向合并的起终点偏移量，横向合并的数目为该项的children数组中所有节点的叶节点总数。
		然后对 向下的剩余部分（绿框区域）递归操作。

 		3.最终得到表头区域数据
			![在这里插入图片描述](https://img-blog.csdnimg.cn/36e001b9b6a440d89f5ecf01d88b133a.png#pic_left)
			
		
		**合并excel单元格的数据描述**
			![在这里插入图片描述](https://img-blog.csdnimg.cn/8242c038d0964a67bb63aa5636391d86.png#pic_left)
3. 其余部分就是填入数据调api即可，可以参考[使用xlsx.js导出有复杂表头的excel](https://blog.csdn.net/seeflyliu/article/details/109476804)这篇文章，此处空白太小所以不再赘述。

# 如果该内容对你有帮助，帮忙star一下项目呀
