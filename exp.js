import XLSX from 'xlsx';
exports.install = function (Vue, options) {
    /**
     * 选择导出类型
     */
    Vue.prototype.selectOutType=function(list){
        let expType =localStorage.getItem(window.sessionStorage.getItem('userid')+'expWaring');
        switch (expType) {
            case '.xlsx':
                this.outXlsx(list,'.xlsx');
                break;
            case '.xls':
                this.outXlsx(list,'.xlsx');
                break;
            case '.csv':
                this.outCsv(list);
                break;
            default:
                this.showConfirm(list)
                break;
        }

    };
    Vue.prototype.showConfirm =function(list){
        let values = '';
        let checkValue='';
        let labels = [{id: '.xlsx'}, {id: '.xls'}, {id: '.csv'}];
        this.$Modal.confirm({
            render: (h) => {
                return h('div', [
                    h('h4', {
                            props: {
                                size: 'small'

                            },
                            style: {
                                color: '#2d8cf0',
                                marginBottom: '10px'
                            }

                        }, '请选择导出格式'
                    ),
                    h('RadioGroup', {
                            props: {
                                value: values
                            },
                            style: {
                                color: '#2d8cf0'
                            },
                            on: {
                                'on-change': value => {
                                    values = value
                                }
                            }
                        },
                        labels.map(p => {
                            return h('Radio', {props: {label: p.id}}, p.id);
                        })),
                    h('Checkbox', {
                            props: {
                                size: 'small',
                                trueValue:true,
                                falseValue:false,
                                value:checkValue
                            },
                            style: {
                                display: 'block',
                                marginTop: '30px',
                                color: '#C0C0C0',
                                fontSize: '12px'
                            },
                            on: {
                                'on-change': value => {
                                    checkValue = value
                                }
                            }

                        }, '记住并且下次不再提醒'
                    )],
                );
            },
            onOk: (e) => {
                if(checkValue===true){
                    localStorage.setItem(window.sessionStorage.getItem('userid')+'expWaring', values)
                }
                switch (values) {
                    case ".xlsx":
                        this.outXlsx(list,'.xlsx')
                        break;
                    case ".xls":
                        this.outXlsx(list,'.xls')
                        break;
                    case ".csv":
                        this.outCsv(list)
                        break;

                }

            },

        })
    };
    Vue.prototype.outXlsx = function (list,name) {
        // if(list.data){
        //     this.formatExcel(list,list.data,k,obj,objarr,objdata)
        // }else{
            this.searchAllData(list,name)
        // }
    };
    Vue.prototype.outCsv = function (list) {
        Promise.all(list.map((i, index) => {
            if(i.data){
                return i.data
            }else{
                return i.fun(...i.params);
            }
        })).then(res => {
            for(let i in res){
                this.expexp(list[i].columns, list[i].ref, list[i].format,res[i])
            }

        })
    };
    /**
     * 导出多页面
     */
    Vue.prototype.searchAllData = function (list,name) {
        let columnHeaders = {};
        let datas = [];

        Promise.all(list.map((i, index) => {
            if(i.data){
                return i.data
            }else{
                return i.fun(...i.params);
            }
        })).then(res => {
            let obj = {};
            let objdata = {};
            let objarr = [];
            for (let k in res) {
                this.formatExcel(list,res,k,obj,objarr,objdata)
            }
            datas = objdata;
            columnHeaders = obj;
            this.outputXlsxFile(
                datas,
                objarr,
                this.$route.meta.title,
                columnHeaders,
                name
            );
        });
    };
    Vue.prototype.formatExcel = function(list,res,k,obj,objarr,objdata){
        res[k] = this.dataDispose(res[k], list[k].format);
        let columns = this.$refs[list[k].name].columns;
        const arr = columns.filter(item => !!item.render);
        obj['detail' + k] = [];
        columns.forEach(i => {
            if (i.key) {
                objarr.push({wch: 20});
                obj['detail' + k].push(i.title);
            }
        });
        // obj['detail' + k] = column.filter(i => i.key);
        let data = res[k];
        objdata['detail' + k] = [];
        data.forEach((item, index) => {
            item.index = index + 1;
            let forobj = {};
            for (var i in arr) {
                let par = {row: {}};
                par.row[arr[i].key] = item[arr[i].key];
                arr[i].render((a, b, c) => {
                    switch (a) {
                        case 'InputNum':
                            item[arr[i].key] = b.props.value;
                            break;
                        case 'Select':
                        case 'Radio' :
                            item[arr[i].key] = obj[arr[i].key].find(item => item.id == b.props.value).value;
                            break;
                        case 'Option':
                            if (obj[arr[i].key] === undefined)obj[arr[i].key] = [];
                            obj[arr[i].key].push({id: b.props.value, value: c});
                            break;
                        default:
                            if (typeof c === 'string' && !c.endsWith('...'))item[arr[i].key] = c;
                            break;
                    }
                }, par);
                forobj[arr[i].title] = item[arr[i].key];
            }

            objdata['detail' + k].push(forobj);
        });
    },
    /**
     * 导出多sheet页
     */
    Vue.prototype.outputXlsxFile = function (data, wscols, xlsxName, columnHeaders,name) {
        let sheetNames = [];
        let sheetsList = {};
        const wb = XLSX.utils.book_new();
        for (let key in data) {
            sheetNames.push(key);
            let columnHeader = columnHeaders[key]; // 此处是每个sheet的表头
            let temp = this.transferData(data[key], columnHeader);
            sheetsList[key] = XLSX.utils.aoa_to_sheet(temp);
            sheetsList[key]['!cols'] = wscols;
        }
        wb['SheetNames'] = sheetNames;
        wb['Sheets'] = sheetsList;
        XLSX.writeFile(wb, xlsxName + name);
    };
    Vue.prototype.transferData = function (data, columnHeader) {
        let content = [];
        content.push(columnHeader);
        data.forEach((item, index) => {
            let arr = [];
            columnHeader.map(column => {
                arr.push(item[column]);
            });
            content.push(arr);
        });
        return content;
    };
    /**
     * 导出逗号转义
     */
    Vue.prototype.expCsv = function (k) {
        for (let m in k) {
            if (typeof k[m] === 'string' && k[m].indexOf(',') !== -1) {
                // 如果还有双引号，先将双引号转义，避免两边加了双引号后转义错误
                if (k[m].indexOf(',') !== -1) {
                    k[m] = k[m].replace('"', '""');
                }
                // 将逗号转义
                k[m] = '"' + k[m] + '"';
            }
        }
        return k;
    };
    Vue.prototype.expexp = function (columns, tableRef, format,data) {
        // 导出基类 span类型没问题 ，如果是render框，则类似select修改，radio同理，如果以后遇到继续加上即可
                const arr = columns.filter(item => !!item.render && !!item.key);
                data = this.dataDispose(data, format);
                let tips = data.map((k, index) => {
                    k.rowno = index + 1;
                    for (var i in arr) {
                        let par = {row: {}};
                        par.row[arr[i].key] = k[arr[i].key];
                        let obj = {};
                        arr[i].render((a, b, c) => {
                            switch (a) {
                                case 'InputNum':
                                    k[arr[i].key] = b.props.value;
                                    break;
                                case 'DateText':
                                    k[arr[i].key] = this.formatss(c);
                                    break;
                                case 'Select':
                                    k[arr[i].key] = obj[arr[i].key].find(item => item.id === b.props.value).value;
                                    break;
                                case 'Option':
                                    if (obj[arr[i].key] === undefined)obj[arr[i].key] = [];
                                    obj[arr[i].key].push({id: b.props.value, value: c});
                                    break;
                                default:
                                    if (typeof c === 'string' && !c.endsWith('...'))k[arr[i].key] = c+'\t';
                                    break;
                            };
                            // 避免类似fixed会渲染两次导致赋值被覆盖
                        }, par);
                    }
                    return this.expCsv(k);
                });
                tableRef.exportCsv({
                    filename: this.$route.meta.title,
                    columns: columns.filter(i => i.type !== 'selection' && i.key !== 'index' && i.type !== 'No'),
                });
    };
};
