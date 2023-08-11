import XLSX from "xlsx-js-style";

/**
 * 一行代码在前端导出excel
 * 会自动处理复杂头和可展开行中的children数据
 * 使用方式：
 *      Excel.export(columns, dataSource, "文件名")  // 前两个入参为 a-table 组件的同名属性值
 * @author dxy
 */
let Excel = {
    /**
     * @param columns       使用 ant table 组件时的 columns 数据
     * @param dataSource    使用 ant table 组件时的 data-source 数据
     * @param fileName      excel导出时的文件名
     */
    export(columns, dataSource, fileName){
        let columnHeight = this.columnHeight(columns)
        let columnWidth = this.columnWidth(columns)
        let header = []
        for(let rowNum =0 ; rowNum < columnHeight; rowNum++){
            header[rowNum] = [];
            for(let colNum =0; colNum < columnWidth; colNum++){
                header[rowNum][colNum] = '';
            }
        }
        let offset = 0;
        let mergeRecord = [];
        for(let item of columns){
            this.generateExcelColumn(header, 0,offset,item,mergeRecord)
            offset += this.treeWidth(item)
        }
        header.push(...this.jsonDataToArray(columns, dataSource))
        let ws = this.aoa_to_sheet(header, columnHeight)
        ws['!merges'] = mergeRecord;
        // 头部冻结
        ws["!freeze"] = {
            xSplit: "1",
            ySplit: "" + columnHeight,
            topLeftCell: "B" + (columnHeight + 1),
            activePane: "bottomRight",
            state: "frozen",
        };
        // 列宽
        ws['!cols'] = [{wpx:165}];
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "sheet1");
        XLSX.writeFile(wb, fileName + ".xlsx");
    },
    aoa_to_sheet(data, headerRows) {
        const ws = {};
        const range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
        // 遍历步骤1里面的二维数组数据
        for (let R = 0; R !== data.length; ++R) {
            for (let C = 0; C !== data[R].length; ++C) {
                if (range.s.r > R) { range.s.r = R; }
                if (range.s.c > C) { range.s.c = C; }
                if (range.e.r < R) { range.e.r = R; }
                if (range.e.c < C) { range.e.c = C; }
                /// 构造cell对象，对所有excel单元格使用如下样式
                let cell;
                if(typeof data[R][C] === "object"){ // 此处预留了自定义设置样式的功能，通过重写recursiveChildrenData方法，可为每一个单元格传入样式属性
                    cell = data[R][C]
                }else{
                    cell = {
                        v: data[R][C],
                        s: {
                            font: { name: "宋体", sz: 11, color: { auto: 1 } },
                            // 单元格对齐方式
                            alignment: {
                                /// 自动换行
                                wrapText: 1,
                                // 水平居中
                                horizontal: "center",
                                // 垂直居中
                                vertical: "center"
                            }
                        }
                    };
                }
                // 头部列表加边框
                if (R < headerRows) {
                    cell.s.border = {
                        top: { style: 'thin', color: { rgb: "000000" } },
                        left: { style: 'thin', color: { rgb: "000000" } },
                        bottom: { style: 'thin', color: { rgb: "000000" } },
                        right: { style: 'thin', color: { rgb: "000000" } },
                    };
                    // 给个背景色
                    cell.s.fill = {
                        patternType: 'solid',
                        fgColor: { theme: 3, "tint": 0.3999755851924192, rgb: 'DDD9C4' },
                        bgColor: { theme: 7, "tint": 0.3999755851924192, rgb: '8064A2' }
                    }
                }
                const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
                // 该单元格的数据类型，只判断了数值类型、布尔类型，字符串类型，省略了其他类型
                // 自己可以翻文档加其他类型
                if (typeof cell.v === 'number') {
                    cell.t = 'n';
                } else if (typeof cell.v === 'boolean') {
                    cell.t = 'b';
                } else {
                    cell.t = 's';
                }
                ws[cell_ref] = cell;
            }
        }
        if (range.s.c < 10000000) {
            ws['!ref'] = XLSX.utils.encode_range(range);
        }
        return ws;
    },
    generateExcelColumn(columnTable, rowOffset, colOffset, columnDefine, mergeRecord){
        let columnWidth = this.treeWidth(columnDefine)
        columnTable[rowOffset][colOffset] = columnDefine.title;
        if(columnDefine.children){
            mergeRecord.push({s:{r:rowOffset, c:colOffset}, e:{r:rowOffset, c:colOffset + columnWidth -1}})
            let tempOffSet = colOffset
            for(let child of columnDefine.children){
                this.generateExcelColumn(columnTable, rowOffset+1, tempOffSet, child,mergeRecord)
                tempOffSet += this.treeWidth(child)
            }
        }else{
            if(rowOffset !== columnTable.length -1)
                mergeRecord.push({s:{r:rowOffset, c:colOffset}, e:{r:columnTable.length -1, c:colOffset}})
        }
    },
    columnHeight(column){
        let height = 0
        for(let item of column){
            height = Math.max(this.treeHeight(item), height)
        }
        return height;
    },
    columnWidth(column){
        let width = 0
        for(let item of column){
            width +=this.treeWidth(item)
        }
        return width;
    },
    treeHeight(root){
        if(root){
            if(root.children && root.children.length!==0){
                let maxChildrenLen = 0;
                for(let child of root.children){
                    maxChildrenLen = Math.max(maxChildrenLen, this.treeHeight(child))
                }
                return  1 + maxChildrenLen;
            }else {
                return 1;
            }
        }else{
            return 0;
        }
    },
    treeWidth(root){
        if(!root) return 0;
        if(!root.children || root.children.length === 0) return 1;
        let width = 0;
        for(let child of root.children){
            width += this.treeWidth(child)
        }
        return width;
    },
    jsonDataToArray(column, data){
        let dataIndexes = [];
        for(let item of column){
            dataIndexes.push(...this.getLeafDataIndexes(item))
        }
        return this.recursiveChildrenData(dataIndexes, data)
    },
    recursiveChildrenData(columnIndex, data){
        let result = [];
        for(let rowData of data){
            let row = [];
            for(let index of columnIndex){
                row.push(rowData[index])
            }
            result.push(row)
            if(rowData.children){
                result.push(...this.recursiveChildrenData(columnIndex, rowData.children))
            }
        }
        return result;
    },
    getLeafDataIndexes(root){
        let result = [];
        if(root.children){
            for(let child of root.children){
                result.push(...this.getLeafDataIndexes(child))
            }
        }else{
            result.push(root.dataIndex);
        }
        return result;
    }
}
export default Excel
