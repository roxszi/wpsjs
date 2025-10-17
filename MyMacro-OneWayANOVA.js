/**
 * 单因素方差分析
 * 1.  数据要分为n个数组
 * 2.  计算每个组的均值、方差、样本数等统计量
 * 3.  计算组间平方和（SSB）、组内平方和（SSW）和总平方和（SST）
 * 4.  计算组间均方（MSB）和组内均方（MSW）
 * 5.  计算F统计量 = MSB / MSW
 * 6.  根据F分布和自由度（组间自由度df1=组数-1，组内自由度df2=总样本数-组数）计算p值
 */
function MY_ONE_WAY_ANOVA() {
  // 起始数据行
  const INIT_ROW = 2
  // 数据组数
  const GROUP_COUNT = 3
  // 分组-大组列
  const BIG_GROUP_COL_1 = "D"
  const BIG_GROUP_COL_2 = "E"
  // 分组-小组列
  const SMALL_GROUP_COL = "F"
  // 数据列
  const DATA_COL = "I"
  // 结果列
  const RESULT_COL = "S"

  // 接表格、表格总行数、函数根对象
  const Sheet = Application.ActiveWorkbook.ActiveSheet
  const rowCount = Sheet.UsedRange.Rows.Count
  const Func = Application.WorksheetFunction

  // 位置索引AOA数组
  const dataRowIndexAoa = [[INIT_ROW]]
  // 数据AOA数组
  const dataAoa = [[Sheet.Range(`${ DATA_COL }${ INIT_ROW }`).Value()]]
  // 组序号所以
  let groupIndex = 0
  // 用于比较的大组名和小组名
  let bigGroupName1 = Sheet.Range(`${ BIG_GROUP_COL_1 }${ INIT_ROW }`).Value()
  let bigGroupName2 = Sheet.Range(`${ BIG_GROUP_COL_2 }${ INIT_ROW }`).Value()
  let smallGroupName = Sheet.Range(`${ SMALL_GROUP_COL }${ INIT_ROW }`).Value()
  // 处理数据
  for (let row = INIT_ROW + 1; row <= rowCount; row++) {
    // 接新大组名和新小组名
    const newBigGroupName1 = Sheet.Range(`${ BIG_GROUP_COL_1 }${ row }`).Value()
    const newBigGroupName2 = Sheet.Range(`${ BIG_GROUP_COL_2 }${ row }`).Value()
    const newSmallGroupName = Sheet.Range(`${ SMALL_GROUP_COL }${ row }`).Value()
    const newData = Sheet.Range(`${ DATA_COL }${ row }`).Value()
    // 如果大组名相同，则比较小组名
    if ((bigGroupName1 === newBigGroupName1) && (bigGroupName2 === newBigGroupName2)) {
      // 如果小组名也相同，则继续添加数据
      if (smallGroupName === newSmallGroupName) {
        dataAoa[groupIndex].push(newData)
      // 如果小组名不同，则说明进入了新小组
      } else {
        // 处理小组位置索引
        dataRowIndexAoa[groupIndex].push(row - 1)
        // 新增小组
        groupIndex++
        dataAoa.push([])
        dataAoa[groupIndex].push(newData)
        // 处理小组位置索引
        dataRowIndexAoa.push([])
        dataRowIndexAoa[groupIndex].push(row)
        // 更新小组名
        smallGroupName = newSmallGroupName
      }
    // 如果大组名不同，则说明进入了新大组
    } else {
      // 处理小组位置索引
      dataRowIndexAoa[groupIndex].push(row - 1)
      // 处理大组数据
      uniProgress(dataAoa, dataRowIndexAoa)
      // 初始化
      dataAoa.length = 0
      dataRowIndexAoa.length = 0
      groupIndex = 0
      // 新增小组
      dataAoa.push([])
      dataAoa[groupIndex].push(newData)
      // 处理小组位置索引
      dataRowIndexAoa.push([])
      dataRowIndexAoa[groupIndex].push(row)
      // 更新小组名和大组名
      bigGroupName1 = newBigGroupName1
      bigGroupName2 = newBigGroupName2
      smallGroupName = newSmallGroupName
    }
  }
  // 处理最后一组数据
  dataRowIndexAoa[groupIndex].push(rowCount)
  uniProgress(dataAoa, dataRowIndexAoa)

  /**
   * 单次处理数据
   * @param { Number[][] } dataAoa 数据数组
   * @param { Number[][] } dataRowIndexAoa 数据位置索引数组
   */
  function uniProgress(dataAoa, dataRowIndexAoa) {
    const groupA = dataAoa[0]
    const groupB = dataAoa[1]
    const groupC = dataAoa[2]
    // 方差齐性检验
    const fTestAB = Func.FTest(groupA, groupB)
    const fTestAC = Func.FTest(groupA, groupC)
    const fTestBC = Func.FTest(groupB, groupC)
    // 总体均值
    const totalMean = Func.Average([...groupA, ...groupB, ...groupC])
    // 组间平方和 (SSB)
    const ssb =
      groupA.length * Math.pow(Func.Average(groupA) - totalMean, 2)
        + groupB.length * Math.pow(Func.Average(groupB) - totalMean, 2)
        + groupC.length * Math.pow(Func.Average(groupC) - totalMean, 2)
    // 组内平方和 (SSW)
    let ssw = 0
    for (const group of [groupA, groupB, groupC]) {
      const groupMean = Func.Average(group)
      for (const uniValue of group) {
        ssw = ssw + Math.pow(uniValue - groupMean, 2)
      }
    }
    // F统计量
    // 组间自由度 = 组数 - 1
    const dfBetween = dataAoa.length - 1
    // 组内自由度 = 总样本数 - 组数
    const dfWithin = groupA.length + groupB.length + groupC.length - dataAoa.length
    // 组间均方（MSB） = 组间平方和 / 组间自由度
    const msb = ssb / dfBetween
    // 组内均方（MSW） = 组内平方和 / 组内自由度
    const msw = ssw / dfWithin
    // F统计量 = MSB / MSW
    const F = msb / msw
    // 获取概率
    const Fp = Func.FDist(F, dfBetween, dfWithin)

    // 输出结果到表格
    Sheet.Range(`${ RESULT_COL }${ dataRowIndexAoa[0][0] }`).Value2 = Fp
    // 合并单元格
    Sheet.Range(`${ RESULT_COL }${ dataRowIndexAoa[0][0] }:${ RESULT_COL }${ dataRowIndexAoa[2][1] }`).Merge(false)

    // 方差齐性检验（带着看一下）
    Sheet.Range(`T${ dataRowIndexAoa[0][0] }`).Value2 = fTestAB
    Sheet.Range(`T${ dataRowIndexAoa[0][0] }:T${ dataRowIndexAoa[2][1] }`).Merge(false)
    Sheet.Range(`U${ dataRowIndexAoa[0][0] }`).Value2 = fTestAC
    Sheet.Range(`U${ dataRowIndexAoa[0][0] }:U${ dataRowIndexAoa[2][1] }`).Merge(false)
    Sheet.Range(`V${ dataRowIndexAoa[0][0] }`).Value2 = fTestBC
    Sheet.Range(`V${ dataRowIndexAoa[0][0] }:V${ dataRowIndexAoa[2][1] }`).Merge(false)

  }

}
