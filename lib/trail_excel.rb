#! ruby -EWindows-31J
# -*- mode:ruby; coding: Windows-31J -*-

require "trail_excel/version"



require 'win32ole'


##----- Excel module -------------------------------

#Authors:: Mt.Trail
#Version:: 1.0 2016/7/17 Mt.Trail
#Copyright:: Copyrigth (C) Mt.Trail 2016 All rights reserved.
#License:: GPL version 2

#= Excel ���p�̂��߂̊g�����W���[��
#==�ړI
# ole32�𗘗p����Excel�𑀍삷��B
#

module Worksheet
  
  #=== �Z���Q��
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
  
  #=== �Z�����
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
  
  #=== �Z���̔w�i�F �Q��
  # sheet.color(y,x)
  #
  def color(y,x)
      self.Cells.Item(y,x).interior.colorindex
  end
  
  #=== �Z���̔w�i�F �ݒ�
  # sheet.color(y,x,color)
  #
  def set_color(y,x,color)
      self.Cells.Item(y,x).interior.colorindex = color
  end
  
  #=== �Z���͈̔͂ւ̔w�i�F �ݒ�
  # set_range_color(y1,x1,y2,x2,color)
  #
  def set_range_color(y1,x1,y2,x2,color)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).interior.colorindex = color
  end
  
  #=== �Z���̕����F �Q��
  # sheet.font_color(y,x)
  #
  def font_color(y,x)
      self.Cells.Item(y,x).Font.colorindex
  end
  
  #=== �Z���̕����F �ݒ�
  # sheet.set_font_color(y,x,color)
  #
  def set_font_color(y,x,color)
      self.Cells.Item(y,x).Font.colorindex = color
  end
  
  #=== �Z���͈̔͂ւ̔w�i�F �ݒ�
  # set_range_font_color(y1,x1,y2,x2,color)
  #
  def set_range_font_color(y1,x1,y2,x2,color)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).Font.colorindex = color
  end
  
  #=== �J�������ݒ�
  #
  def set_width(y,x,width)
      self.Cells.Item(y,x).ColumnWidth = width
  end
  
  #=== �s�̍����ݒ�
  #
  def set_height(y,x,height)
      self.Cells.Item(y,x).RowHeight = height
  end


  #=== �Z���ʒu���w�肷�镶����쐬
  #
  def r_str(y,x)
    self.Cells.Item(y,x).address('RowAbsolute'=>false,'ColumnAbsolute'=>false)
  end
  
  #=== �Z����I��
  #
  def select( y,x)
    r = r_str(y,x)
    self.Range(r).select
  end

  #=== �Z���Ɏ���ݒ�
  #
  def formula( y,x,f)
    r = r_str(y,x)
    self.Range(r).Formula = f
  end
  
  #=== �Z���ɐݒ肳�ꂽ�����Q��
  #
  def get_formula( y,x)
    r = r_str(y,x)
    self.Range(r).Formula
  end
  
  #=== �s�̃O���[�v��
  #
  def group_row(y1,y2)
    r = r_str(y1,1)+':'+r_str(y2,1)
    self.Range(r).Rows.Group
  end
  
  #=== �J�����̃O���[�v��
  #
  def group_column(x1,x2)
    r = r_str(1,x1)+':'+r_str(1,x2)
    self.Range(r).Columns.Group
  end
  
  #=== �w��͈͂��}�[�W
  #
  def merge(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).MergeCells = true
  end
  
  #=== �g����ݒ�
  #
  def box(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).Borders.LineStyle = 1
  end
  
  #=== ������̐܂�Ԃ��w��
  #
  def wrap(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).HorizontalAlignment = 1
    self.Range(r).WrapText = true
  end
  
  #=== ��t��
  #
  def v_top(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).VerticalAlignment = -4160
  end
  
  #=== ��������
  #
  def center(y1,x1,y2,x2)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r).HorizontalAlignment = -4108
  end
  
  #=== �͈͎w��̎��̃R�s�[
  #
  def format_copy(y1,x1,y2,x2,y3,x3)
    r2 = r_str(y3,x3)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r2).Copy
    self.Range(r).PasteSpecial('Paste' => -4122)
  end
  
  #=== �Z���̎��̃R�s�[
  #
  def format_copy1(y1,x1,y2,x2)
    r2 = r_str(y2,x2)
    r = r_str(y1,x1)
    self.Range(r2).Copy
    self.Range(r).PasteSpecial('Paste' => -4122)
  end
  
  #=== �R�s�[
  #
  def copy(y1,x1,y2,x2,y3,x3)
    r2 = r_str(y3,x3)
    r = r_str(y1,x1)+':'+r_str(y2,x2)
    self.Range(r2).Copy
    self.Range(r).PasteSpecial('Paste' => -4104)
  end
  
  #=== �s�̑}��
  #
  def insert_row(n)
    self.Rows("#{n}:#{n}").Insert('Shift' => -4121)
  end
  
  #=== �s�̍폜
  #
  def delete_row(n,m)
    self.Range("#{n}:#{m}").Delete
  end

  #=== �摜�̓\��t���@point�ʒu�w��
  #
  def add_picture(py,px,file,sh,sw)
    self.Shapes.AddPicture(file,false,true,px,py,0.75*sw,0.75*sh)
  end

  #=== �摜�̓\��t���@�Z���ʒu�w��
  #
  def add_picture_at_cell(cy,cx,file,sh,sw)
    r = self.Range(r_str(cy,cx))
    self.Shapes.AddPicture(file,false,true,r.Left,r.Top,0.75*sw,0.75*sh)
  end

end

##----- End of Excel module -------------------------------

#=== ��΃p�X��
#
def getAbsolutePath filename
  fso = WIN32OLE.new('Scripting.FileSystemObject')
  fn = fso.GetAbsolutePathName(filename).gsub('\\','/')
  fn
end

#=== Excel�̃I�[�v��
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

#=== Excel�̃u�b�N����
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

