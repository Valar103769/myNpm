import type { Worksheet } from 'exceljs'
import {
  addHeaderStyle,
  DEFAULT_COLUMN_WIDTH,
  generateHeaders,
  getColumnNumber,
  mergeColumnCell,
  mergeRowCell,
  saveWorkbook,
} from './shared'
import type { ITableHeader } from './types'

type Props = {
  dataSource: any[]
  columns: any[]
  formData: Record<'filename', string>
}

export async function onExportMultiHeaderExcel(props: Props) {
  const {
    dataSource,
    columns,
    formData: { filename = 'excel' },
  } = props

  const ExcelJs = await import('exceljs')
  // 创建工作簿
  const workbook = new ExcelJs.Workbook()
  // 添加sheet
  const worksheet = workbook.addWorksheet('sheet')
  // 设置 sheet 的默认行高
  worksheet.properties.defaultRowHeight = 20
  // 解析 AntD Table 的 columns
  const headers = generateHeaders(columns)
  console.log({ headers })
  // 第一行表头
  const names1: string[] = []
  // 第二行表头
  const names2: string[] = []
  // 用于匹配数据的 keys
  const headerKeys: string[] = []
  headers.forEach((item) => {
    if (item.children) {
      // 有 children 说明是多级表头，header name 需要两行
      item.children.forEach((child) => {
        names1.push(item.header)
        names2.push(child.header)
        headerKeys.push(child.key)
      })
    } else {
      const columnNumber = getColumnNumber(item.width)
      for (let i = 0; i < columnNumber; i++) {
        names1.push(item.header)
        names2.push(item.header)
        headerKeys.push(item.key)
      }
    }
  })
  handleHeader(worksheet, headers, names1, names2)
  // 添加数据
  addData2Table(worksheet, headerKeys, headers, dataSource)
  // 给每列设置固定宽度
  worksheet.columns = worksheet.columns.map((col) => ({ ...col, width: DEFAULT_COLUMN_WIDTH }))
  // 导出excel
  saveWorkbook(workbook, `${filename}.xlsx`)
}

function handleHeader(
  worksheet: Worksheet,
  headers: ITableHeader[],
  names1: string[],
  names2: string[],
) {
  // 判断是否有 children, 有的话是两行表头
  const isMultiHeader = headers?.some((item) => item.children)
  if (isMultiHeader) {
    // 加表头数据
    const rowHeader1 = worksheet.addRow(names1)
    const rowHeader2 = worksheet.addRow(names2)
    // 添加表头样式
    addHeaderStyle(rowHeader1, { color: 'dff8ff' })
    addHeaderStyle(rowHeader2, { color: 'dff8ff' })
    mergeColumnCell(headers, rowHeader1, rowHeader2, names1, names2, worksheet)
    return
  }
  // 加表头数据
  const rowHeader = worksheet.addRow(names1)
  // 表头根据内容宽度合并单元格
  mergeRowCell(headers, rowHeader, worksheet)
  // 添加表头样式
  addHeaderStyle(rowHeader, { color: 'dff8ff' })
}

function addData2Table(
  worksheet: Worksheet,
  headerKeys: string[],
  headers: ITableHeader[],
  dataSource: any[],
) {
  dataSource?.forEach((item: any) => {
    const rowData = headerKeys?.map((key) => item[key])
    const row = worksheet.addRow(rowData)
    mergeRowCell(headers, row, worksheet)
    row.height = 26
    // 设置行样式, wrapText: 自动换行
    row.alignment = { vertical: 'middle', wrapText: true, shrinkToFit: false }
    row.font = { size: 11, name: '微软雅黑' }
  })
}
