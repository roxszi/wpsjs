/**
 * 把小组里的数据优选一下，然后算均值
 * 1.  把3或4个数据，先算均值
 * 2.  然后计算每个数据与均值的差值
 * 3.  把差值大的剔除，只保留差值小的2个数据
 * 4.  然后再算均值
 */
function MY_PREFERRED_AVERAGE() {
  // 起始数据行
  const INIT_ROW = 2
  // 分组-大组列
  const BIG_GROUP_COL_1 = "D"
  const BIG_GROUP_COL_2 = "E"
  // 分组-小组列
  const SMALL_GROUP_COL = "F"
  // 数据列
  const DATA_COL = "I"
  // 结果列
  const RESULT_COL = "K"

  // 接表格、表格总行数、函数根对象
  const Sheet = Application.ActiveWorkbook.ActiveSheet
  const rowCount = Sheet.UsedRange.Rows.Count
  const Func = Application.WorksheetFunction

  // 位置索引数组
  const dataRowIndexArr = [INIT_ROW]
  // 数据数组
  const dataArr = [Sheet.Range(`${ DATA_COL }${ INIT_ROW }`).Value()]
  // 用于比较的组名
  let bigGroupName1 = Sheet.Range(`${ BIG_GROUP_COL_1 }${ INIT_ROW }`).Value()
  let bigGroupName2 = Sheet.Range(`${ BIG_GROUP_COL_2 }${ INIT_ROW }`).Value()
  // 处理数据
  for (let row = INIT_ROW + 1; row <= rowCount; row++) {
    // 接新组名
    const newBigGroupName1 = Sheet.Range(`${ BIG_GROUP_COL_1 }${ row }`).Value()
    const newBigGroupName2 = Sheet.Range(`${ BIG_GROUP_COL_2 }${ row }`).Value()
    const newData = Sheet.Range(`${ DATA_COL }${ row }`).Value()
    // 如果组名相同，则继续添加数据
    if ((bigGroupName1 === newBigGroupName1) && (bigGroupName2 === newBigGroupName2)) {
      dataArr.push(newData)
    // 如果组名不同，则说明进入了新组
    } else {
      // 处理位置索引
      dataRowIndexArr.push(row - 1)
      uniProgress(dataArr, dataRowIndexArr)
      // 初始化
      dataArr.length = 0
      dataRowIndexArr.length = 0
      // 新增数据、处理位置索引
      dataArr.push(newData)
      dataRowIndexArr.push(row)
      // 更新组名
      bigGroupName1 = newBigGroupName1
      bigGroupName2 = newBigGroupName2
    }
  }
  // 处理最后一组数据
  dataRowIndexArr.push(rowCount)
  uniProgress(dataArr, dataRowIndexArr)

  /**
   * 单次处理数据
   * @param { Number[] } dataArr 数据数组
   * @param { Number[] } dataRowIndexArr 数据位置索引数组
   */
  function uniProgress(dataArr, dataRowIndexArr) {
    // 均值
    const ave = Func.Average(dataArr)
    // 计算差值
    const diffArr = []
    // 排序准备
    for (const item of dataArr) {
      diffArr.push([item, Math.abs(item - ave)])
    }
    // 按差值排序
    diffArr.sort((a, b) => a[1] - b[1])
    // 优选2个数据
    const preferredDataArr = []
    for (const item of diffArr.slice(0, 2)) {
      preferredDataArr.push(item[0])
    }
    // 优选均值
    const preferredAve = Func.Average(preferredDataArr)

    // 输出结果到表格
    Sheet.Range(`${ RESULT_COL }${ dataRowIndexArr[0] }`).Value2 = preferredAve
    // 合并单元格
    Sheet.Range(`${ RESULT_COL }${ dataRowIndexArr[0] }:${ RESULT_COL }${ dataRowIndexArr[1] }`).Merge(false)

  }

}
