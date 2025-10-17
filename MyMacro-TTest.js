/**
 * 计算两组数据的均值、标准差、T检验P值
 */
function MY_TTEST(index) {
  // 初始化起始行数
  const FIRST_INDEX = 2
  // 接表格、表格总行数
  const Sheet = Application.ActiveWorkbook.ActiveSheet
  const rowCount = Sheet.UsedRange.Rows.Count
  // 建个记录序列的数组，把第1个值初始化进去
  const rangeIndexArr = [FIRST_INDEX]
  // “文件名”和“处理人”，并初始化
  let fileName = Sheet.Range(`G${ FIRST_INDEX }`).Value()
  let personName = Sheet.Range(`H${ FIRST_INDEX }`).Value()
  // 遍历所有行（从数据行开始）
  for (let i = FIRST_INDEX + 1; i <= rowCount; i++) {
    // 接新“文件名”和新“处理人”
    const newFileName = Sheet.Range(`G${ i }`).Value()
    const newPersonName = Sheet.Range(`H${ i }`).Value()
    // 比较文件名：若文件名相等
    if (fileName === newFileName) {
      // 比较处理人：若处理人不相等
      if (personName !== newPersonName) {
        // 把上一个值和本值推进数组
        rangeIndexArr.push(i - 1, i)
        // 刷新处理人
        personName = newPersonName
      }
    // 若文件名不相等
    } else {
      // 把上一个值推进数组
      rangeIndexArr.push(i - 1)
      // 处理数据：
      uniProgress(rangeIndexArr)
      // 清空数组并推进本值
      rangeIndexArr.length = 0
      rangeIndexArr.push(i)
      // 刷新文件名和处理人
      fileName = newFileName
      personName = newPersonName
    }
  }
  // 最后，把剩下的数据处理一次（也就是最后一组）
  rangeIndexArr.push(rowCount)
  uniProgress(rangeIndexArr)
  // 处理结束，保存
  Application.Save()

  /**
   * 单次处理数据
   * 会计算两组的均值Ave.、标准差SD、T检验P值（会根据F检验结果，自动使用方差齐/不齐的T检验）
   * @param { Number[] } rangeIndexArr 传参分别为2段数据的收尾index
   */
  function uniProgress(rangeIndexArr) {
    // console.log(
    //   "rangeIndexArr: ",
    //   rangeIndexArr[0],
    //   rangeIndexArr[1],
    //   rangeIndexArr[2],
    //   rangeIndexArr[3]
    // )
    // 接函数根对象
    const Func = Application.WorksheetFunction
    // 第一组数据范围
    const range1 = `I${ rangeIndexArr[0] }:I${ rangeIndexArr[1] }`
    // 第二组数据范围
    const range2 = `I${ rangeIndexArr[2] }:I${ rangeIndexArr[3] }`
    // 第一组数据均值、SD，并合并单元格
    Sheet.Range(`O${ rangeIndexArr[0] }`).Formula = `=AVERAGE(${ range1 })`
    Sheet.Range(`P${ rangeIndexArr[0] }`).Formula = `=STDEV(${ range1 })`
    Sheet.Range(`O${ rangeIndexArr[0] }:O${ rangeIndexArr[1] }`).Merge(false)
    Sheet.Range(`P${ rangeIndexArr[0] }:P${ rangeIndexArr[1] }`).Merge(false)
    // 第二组数据均值、SD，并合并单元格
    Sheet.Range(`O${ rangeIndexArr[2] }`).Formula = `=AVERAGE(${ range2 })`
    Sheet.Range(`P${ rangeIndexArr[2] }`).Formula = `=STDEV(${ range2 })`
    Sheet.Range(`O${ rangeIndexArr[2] }:O${ rangeIndexArr[3] }`).Merge(false)
    Sheet.Range(`P${ rangeIndexArr[2] }:P${ rangeIndexArr[3] }`).Merge(false)
    // 均值差，并合并单元格
    Sheet.Range(`Q${ rangeIndexArr[0] }`).Formula = `=ABS(O${ rangeIndexArr[0] }-O${ rangeIndexArr[2] })`
    Sheet.Range(`Q${ rangeIndexArr[0] }:Q${ rangeIndexArr[3] }`).Merge(false)
    // F检验（方差齐性检验）
    const FTestValue = Func.FTest(Sheet.Range(range1).Value(), Sheet.Range(range2).Value())
    // 若F检验的p值小于0.05，则执行T检验（方差不齐）
    if (FTestValue < 0.05) {
      Sheet.Range(`R${ rangeIndexArr[0] }`).Formula = `=TTEST(${ range1 }, ${ range2 }, 2, 3)`
    // 否则，执行T检验（方差齐）
    } else {
      Sheet.Range(`R${ rangeIndexArr[0] }`).Formula = `=TTEST(${ range1 }, ${ range2 }, 2, 2)`
    }
    // 合并
    Sheet.Range(`R${ rangeIndexArr[0] }:R${ rangeIndexArr[3] }`).Merge(false)
  }
}
