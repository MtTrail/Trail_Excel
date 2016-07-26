#! ruby -EWindows-31J
# -*- mode:ruby; coding: Windows-31J -*-

require "trail_excel/version"



require 'win32ole'


##----- Excel module -------------------------------

#Authors:: Mt.Trail
#Version:: 1.0 2016/7/17 Mt.Trail
#Copyright:: Copyrigth (C) Mt.Trail 2016 All rights reserved.
#License:: GPL version 2

#= Excel 利用のための拡張モジュール
#==目的
# ole32を利用してExcelを操作する。
#

module Worksheet
  
  #=== セル参照
  # sheet[y,x]
  #
  def [] y,x
    cell = self.Cells.Item(y,x)
    if cell.MergeCells
      cell.MergeArea.Item(1,1).Value
    else
      cell.Value
    end
  end
  
  #=== セル代入
  # sheet[y,x] = xx
  #
  def []= y,x,value
    cell = self.Cells.Item(y,x)
    if cell.MergeCells
      cell.MergeArea.Item(1,1).Value = value
    else
      cell.Value = value
    end
  end
  
  #=== セルの背景色 参照
  # sheet.color(y,x)
  #
  def color(y,x)
      self.Cells.Item(y,x).interior.colorindex
  end
  
  #=== セルの背景色 設定
  # sheet.color(y,x,color)
  #
  def set_color(y,x,color)
      self.Cells.Item(y,x).interior.colorindex = color
  end
  
  #=== セルの範囲への背景色 設定
  # set_range_color(y1,x1,y2,x2,color)
  #
  def set_range_color(y1,x1,y2,x2,color)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).interior.colorindex = color
  end
  
  #=== セルの文字色 参照
  # sheet.font_color(y,x)
  #
  def font_color(y,x)
      self.Cells.Item(y,x).Font.colorindex
  end
  
  #=== セルの文字色 設定
  # sheet.set_font_color(y,x,color)
  #
  def set_font_color(y,x,color)
      self.Cells.Item(y,x).Font.colorindex = color
  end
  
  #=== セルの範囲への背景色 設定
  # set_range_font_color(y1,x1,y2,x2,color)
  #
  def set_range_font_color(y1,x1,y2,x2,color)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).Font.colorindex = color
  end
  
  #=== カラム幅設定
  #
  def set_width(y,x,width)
      self.Cells.Item(y,x).ColumnWidth = width
  end
  
  #=== 行の高さ設定
  #
  def set_height(y,x,height)
      self.Cells.Item(y,x).RowHeight = height
  end


  #=== セル位置を指定する文字列作成
  #
  def r_str(y,x)
    self.Cells.Item(y,x).address('RowAbsolute'=>false,'ColumnAbsolute'=>false)
  end
  
  #=== セルを選択
  #
  def select( y,x)
    r = r_str(y,x)
    self.Range(r).select
  end

  #=== セルに式を設定
  #
  def formula( y,x,f)
    r = r_str(y,x)
    self.Range(r).Formula = f
  end
  
  #=== セルに設定された式を参照
  #
  def get_formula( y,x)
    r = r_str(y,x)
    self.Range(r).Formula
  end
  
  #=== 行のグループ化
  #
  def group_row(y1,y2)
    r = r_str(y1,1)+':'+r_str(y2,1)
    self.Range(r).Rows.Group
  end
  
  #=== カラムのグループ化
  #
  def group_column(x1,x2)
    r = r_str(1,x1)+':'+r_str(1,x2)
    self.Range(r).Columns.Group
  end
  
  #=== 指定範囲をマージ
  #
  def merge(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).MergeCells = true
  end
  
  #=== 枠線を設定
  #
  def box(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).Borders.LineStyle = 1
  end
  
  #=== 文字列の折り返し指定
  #
  def wrap(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).HorizontalAlignment = 1
    self.Range(r).WrapText = true
  end
  
  #=== 上付き
  #
  def v_top(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).VerticalAlignment = -4160
  end
  
  #=== 中央揃え
  #
  def center(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).HorizontalAlignment = -4108
  end
  
  #=== 範囲指定の式のコピー
  #
  def format_copy(y1,x1,y2,x2,y3,x3)
    r2 = r_str(y3,x3)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r2).Copy
    self.Range(r).PasteSpecial('Paste' => -4122)
  end
  
  #=== セルの式のコピー
  #
  def format_copy1(y1,x1,y2,x2)
    r2 = r_str(y2,x2)
    r = r_str(y1,x1)
    self.Range(r2).Copy
    self.Range(r).PasteSpecial('Paste' => -4122)
  end
  
  #=== コピー
  #
  def copy(y1,x1,y2,x2,y3,x3)
    r2 = r_str(y3,x3)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r2).Copy
    self.Range(r).PasteSpecial('Paste' => -4104)
  end
  
  #=== 行の挿入
  #
  def insert_row(n)
    self.Rows("#{n}:#{n}").Insert('Shift' => -4121)
  end
  
  #=== 行の削除
  #
  def delete_row(n,m)
    self.Range("#{n}:#{m}").Delete
  end

  #=== 画像の貼り付け　point位置指定
  #
  def add_picture(py,px,file,sh,sw)
    self.Shapes.AddPicture(file,false,true,px,py,0.75*sw,0.75*sh)
  end

  #=== 画像の貼り付け　セル位置指定
  #
  def add_picture_at_cell(cy,cx,file,sh,sw)
    r = self.Range(r_str(cy,cx))
    self.Shapes.AddPicture(file,false,true,r.Left,r.Top,0.75*sw,0.75*sh)
  end

end

##----- End of Excel module -------------------------------

#=== 絶対パス化
#
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  fn = fso.GetAbsolutePathName(filename).gsub('\\','/')
  fn
end

#=== Excelのオープン
#
def openExcelWorkbook (filename, visible:false, pw:nil)
  filename = getAbsolutePath(filename)
  xl = WIN32OLE.new('Excel.Application')
  xl.Visible = visible
  xl.DisplayAlerts = true
  if pw
    book = xl.Workbooks.Open(:FileName=>"#{filename}",:Password=>pw)
  else
    book = xl.Workbooks.Open(filename)
  end

  begin
    yield book
  ensure
    xl.Workbooks.Close
    xl.Quit
  end
end

#=== Excelのブック生成
#
def createExcelWorkbook
  xl = WIN32OLE.new('Excel.Application')
  xl.Visible = false
  xl.DisplayAlerts = false
  book = xl.Workbooks.Add()
  begin
    yield book
  ensure
    xl.Workbooks.Close
    xl.Quit
  end
end

