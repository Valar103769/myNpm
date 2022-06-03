/* eslint-disable @typescript-eslint/no-unused-vars */

import { cloneDeep, isEmpty } from 'lodash-es'
import { generateHeaders, saveWorkbook } from './shared'

type Props = {
  dataSource: any[]
  columns: any[]
  formData: Record<'filename', string>
}
// 导出
export async function onExportBasicExcelWithStyle(props: Props) {
  const {
    dataSource,
    columns,
    formData: { filename = 'excel' },
  } = props

  if (isEmpty(dataSource) || isEmpty(columns)) {
    return
  }

  const ExcelJs = await import('exceljs')
  let list = cloneDeep(dataSource)
  // 处理 antd 的render
  columns?.forEach((column: any) => {
    if (column.renderText) {
      list = list.map((row: any, index: number) => ({
        ...row,
        [column.dataIndex]: column.renderText(row[column.dataIndex], row, index),
      }))
    }
  })

  // 创建工作簿
  const workbook = new ExcelJs.Workbook()
  // 添加sheet
  const worksheet = workbook.addWorksheet('sheet')
  // 设置 sheet 的默认行高
  worksheet.properties.defaultRowHeight = 20
  // 设置列
  worksheet.columns = generateHeaders(columns)
  // 给表头添加背景色。因为表头是第一行，可以通过 getRow(1) 来获取表头这一行
  const headerRow = worksheet.getRow(1)
  // 直接给这一行设置背景色
  // headerRow.fill = {
  //   type: 'pattern',
  //   pattern: 'solid',
  //   fgColor: {argb: 'dff8ff'},
  // }
  // 通过 cell 设置样式，更精准
  headerRow.eachCell((cell, colNum) => {
    // 设置背景色
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'dff8ff' },
    }
    // 设置字体
    cell.font = {
      bold: true,
      italic: true,
      size: 12,
      name: '微软雅黑',
      color: { argb: 'ff0000' },
    }
    // 设置对齐方式
    cell.alignment = { vertical: 'middle', horizontal: 'left', wrapText: false }
  })
  // 添加行
  const rows = worksheet.addRows(list)
  // 设置每行的样式
  rows?.forEach((row) => {
    // 设置字体
    row.font = {
      size: 11,
      name: '微软雅黑',
    }
    // 设置对齐方式
    row.alignment = { vertical: 'middle', horizontal: 'left', wrapText: false }
  })
  // 导出excel
  saveWorkbook(workbook, `${filename}.xlsx`)
}
