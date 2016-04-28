# coding: utf-8
DEFAULT_TK_ENCODING=__ENCODING__

require 'win32ole'
require 'tk'

def check_for_open(x,y)
  menu_open_document_clicked if $tree.identify_item(x,y) == $tree.focus_item
end

def print_document(action_window,filename,sum)
  exellio_connect(action_window) {|exellio|
    begin
      exellio.OpenFiscalReceipt(1,"0000",1)
  
      answer = (exellio.LastError==0)?'ok':(Tk::messageBox :parent => action_window, :message => exellio.LastErrorText, :title => "Ошибка создания чека", :icon => 'error', :type => 'retrycancel')
  
      if answer == 'cancel'
        exellio.CancelReceipt()
        next false
      end
    end while answer == 'retry'

    printed = true
  
    file = File.open(filename,"r:cp1251")
    txt = file.readlines
    file.close
  
    txt.each_with_index {|str,index|
      items = str.split(/\s*;\s*/)
  
      if items.length == 6 then
        exellio.Sale(items[0].to_i, items[1], items[2].to_i, items[3].to_i, items[4].to_f, items[5].to_f, 0, 0, false, "0000")
        
        answer = (exellio.LastError==0)?'ok':(Tk::messageBox :parent => action_window, :message => exellio.LastErrorText, :title => "Ошибка печати строки чека", :icon => 'error', :type => 'retrycancel')
      
        if answer == 'cancel'
          exellio.CancelReceipt()
          printed = false
      
          break
        elsif answer == 'retry'
          redo
        end
      else
        exellio.CancelReceipt()

        Tk::messageBox :parent => action_window, :message => 'Неправильное количество полей в файле!', :title => "Ошибка печати строки чека", :icon => 'error', :type => 'retrycancel'

        printed = false

        break
      end
    }

    next false unless printed

    begin
      exellio.TotalEx("",1,sum)

      answer = (exellio.LastError==0)?'ok':(Tk::messageBox :parent => action_window, :message => exellio.LastErrorText, :title => "Ошибка подсчета итога чека", :icon => 'error', :type => 'retrycancel')

      if answer == 'cancel'
        exellio.CancelReceipt()
        next false
      end
    end while answer == 'retry'
    
    true
  }
end

def check_entry(p,dec=0)
  if p.empty? then true
  elsif dec == 0 then 
    p == p[/\d{1,9}/]    
  else
    p == p[/\d{1,9}((\.|,)\d{0,#{dec}})?/]
  end
end

def check_mouse_click(x,y,window, tk_variable)
  (window.destroy; tk_variable.value = -1) if (x < 0)||(x > window.winfo_width)||(y < 0)||(y > window.winfo_height)
end

def check_keypressed(key, window, label_total, parent, tk_variable, dec=0)
  case key
    when 'Escape' then window.destroy; tk_variable.value = -1
    when 'Return' then parent.value = (label_total.text.to_f==0)?'':label_total.text; window.destroy; tk_variable.value = -1
    when 'Delete' then label_total.text = '0'
    when 'BackSpace' then label_total.text = label_total.text[0...label_total.text.length-1]; label_total.text = '0' if label_total.text.empty? 
  else
    new_value = label_total.text+key.sub(/comma|period/,'.')
    if check_entry(new_value,dec)
      if (label_total.text == label_total.text[/0*/])&(key.sub(/comma|period/,'.') != '.') then label_total.text=key.to_i 
      else label_total.text = new_value
      end
    end
  end
end

def calculator(window,parent,tk_variable,dec=0)
  parent.focus
  
  window_calculator = Tk::Toplevel.new(window) {focus; withdraw; resizable false,false; overrideredirect true; grab 'global'}
  
  parent_value = parent.value[/(0(?=(\.|,))|[1-9])\d*(\.|,)?(\d{0,#{dec}})?/]
  
  frame = Tk::Tile::Frame.new(window_calculator){relief 'solid'; borderwidth 3; pack :expand => true, :fill => 'both'}
  label_total = Tk::Tile::Label.new(frame) {text "#{(parent_value)?parent_value:'0'}"; font 'system'; background 'CornflowerBlue'; foreground 'snow2'; anchor 'e'; relief 'sunken'; grid :column=>0, :row=>0, :columnspan => 4, :sticky => 'ew'}
  Tk::Tile::Button.new(frame) {text 1; command proc{check_keypressed('1',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>0, :row=>1, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 2; command proc{check_keypressed('2',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>1, :row=>1, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 3; command proc{check_keypressed('3',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>2, :row=>1, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 'C'; command proc{check_keypressed('Delete',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>3, :row=>1, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 4; command proc{check_keypressed('4',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>0, :row=>2, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 5; command proc{check_keypressed('5',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>1, :row=>2, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 6; command proc{check_keypressed('6',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>2, :row=>2, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text '<-'; command proc{check_keypressed('BackSpace',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>3, :row=>2, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 7; command proc{check_keypressed('7',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>0, :row=>3, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 8; command proc{check_keypressed('8',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>1, :row=>3, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 9; command proc{check_keypressed('9',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>2, :row=>3, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text '00'; command proc{check_keypressed('00',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>0, :row=>4, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 0; command proc{check_keypressed('0',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>1, :row=>4, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text '.'; state "#{(dec)==0?'disabled':'normal'}"; command proc{check_keypressed('.',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>2, :row=>4, :sticky => 'news'}
  Tk::Tile::Button.new(frame) {text 'OK'; command proc{check_keypressed('Return',window_calculator,label_total,parent,tk_variable,dec)}; grid :column=>3, :row=>3, :rowspan => 2, :sticky => 'news'}
  
  frame.grid_columnconfigure(0, :weight=>1, :uniform=>'1')
  frame.grid_columnconfigure(1, :weight=>1, :uniform=>'1')
  frame.grid_columnconfigure(2, :weight=>1, :uniform=>'1')
  frame.grid_columnconfigure(3, :weight=>1, :uniform=>'1')

  frame.grid_rowconfigure(1, :weight=>1, :uniform=>'1')
  frame.grid_rowconfigure(2, :weight=>1, :uniform=>'1')
  frame.grid_rowconfigure(3, :weight=>1, :uniform=>'1')
  frame.grid_rowconfigure(4, :weight=>1, :uniform=>'1')

  window_calculator.bind('ButtonPress', proc{|x,y| check_mouse_click(x,y,window_calculator,tk_variable)},'%x %y') 
  window_calculator.bind('KeyPress', proc{|key| check_keypressed(key,window_calculator,label_total,parent,tk_variable,dec)},'%K')
  
  window_calculator.update

  window_calculator.geometry("150x150+#{parent.winfo_rootx}+#{parent.winfo_rooty+parent.winfo_height}")

  window_calculator.deiconify
end

def calculate(key,value,hash,hash_bills,hash_coins, total_entry)
  return false unless check_entry(value)
  
  hash[key] = value.to_i
  
  sum = 0
  
  hash_bills.each_pair {|key,value| sum += key*value}
  hash_coins.each_pair {|key,value| sum += key*value}
  
  total_entry.validate('none')
  total_entry.value = "%.2f"%sum
  total_entry.validate('key')
  
  true
end

def get_frame_bills_counter(parent,total_entry)
  bills = {'500'=>500,'200'=>200,'100'=>100,'50'=>50,'20'=>20,'10'=>10,'5'=>5,'2'=>2,'1'=>1}
  coins = {'1 грн.'=>1,'50'=>0.5,'25'=>0.25,'10'=>0.10,'5'=>0.05,'2'=>0.02,'1'=>0.01}
  
  bills_count = {}
  coins_count = {}
    
  bills_entrys = []
  coins_entrys = []

  tk_variable = TkVariable.new(-1)
  
  frame_bills_counter = Tk::Tile::Frame.new(parent)
  
  frame_bill = Tk::Tile::LabelFrame.new(frame_bills_counter) {text 'Купюры'; grid :column => 0, :row => 1, :sticky => 'news', :padx => [5,2], :pady => 5}
  
  bills.each_with_index {|(key, value), index|
    bills_count[value] = 0
    
    Tk::Tile::Label.new(frame_bill) {text "#{key}:"; anchor 'e'; grid :column => 0, :row => index, :sticky => 'we', :padx => [5,0], :pady => [0,5]}
    bills_entrys[index] = Tk::Tile::Entry.new(frame_bill) {justify 'right'; width 9; validate 'key'; validatecommand proc{|p| calculate(value,p,bills_count,bills_count,coins_count,total_entry)},'%P'; grid :column => 1, :row => index, :sticky => 'ew', :padx => [0,5], :pady => [0,5], :ipadx => 30}
    Tk::Tile::Checkbutton.new(bills_entrys[index]) {text '...'; takefocus 0; variable tk_variable; onvalue index; offvalue -1; command proc{calculator(parent,bills_entrys[index],tk_variable)}; style 'Toolbutton'; cursor 'arrow'; pack :side => 'left', :padx => 1, :pady => 1}
  }
  frame_bill.grid_columnconfigure(1, :weight => 1)

  frame_coins = Tk::Tile::LabelFrame.new(frame_bills_counter) {text 'Монеты'; grid :column => 1, :row => 1, :sticky => 'news', :padx => [2,5], :pady => 5}
  coins.each_with_index {|(key, value),index|
    coins_count[value] = 0

    Tk::Tile::Label.new(frame_coins) {text "#{key}:"; anchor 'e'; grid :column => 0, :row => index, :sticky => 'we', :padx => [5,0], :pady => [0,5]}
    coins_entrys[index] = Tk::Tile::Entry.new(frame_coins) {justify 'right'; width 9; validate 'key'; validatecommand proc{|p| calculate(value,p,coins_count,bills_count,coins_count,total_entry)},'%P'; grid :column => 1, :row => index, :sticky => 'we', :padx => [0,5], :pady => [0,5], :ipadx => 30}
    Tk::Tile::Checkbutton.new(coins_entrys[index]) {text '...'; takefocus 0; variable tk_variable; onvalue index+bills.length; offvalue -1; style 'Toolbutton'; cursor 'arrow'; command proc{calculator(parent,coins_entrys[index], tk_variable)}; pack :side => 'left', :padx => 1, :pady => 1}
  }
  frame_coins.grid_columnconfigure(1, :weight => 1)
  
  return frame_bills_counter, bills_count, coins_count, bills_entrys, coins_entrys
end

def exellio_connect(parent_window)
  exellio = WIN32OLE.new('ExellioFP.FiscalPrinter')

  begin
    parent_window.cursor = 'watch'
    parent_window.update

    begin
      exellio.OpenPort($comport,115200)

      answer = (exellio.LastError==0)?'ok':(Tk::messageBox :parent => parent_window, :message => exellio.LastErrorText, :title => "Ошибка подключения", :icon => 'error', :type => 'retrycancel')

      return false if answer == 'cancel'
    end while answer == 'retry'

    while exellio.GetStatusBit(1,5) == 1
      return false if (Tk::messageBox :parent => parent_window, :message => 'Открыта крышка фискального принтера.\nЗакройте и нажмите OK для продолжения', :title => "Ошибка подключения", :icon => 'error', :type => 'okcancel')=='cancel'
    end

    if exellio.GetStatusBit(2,0) == 1
      Tk::messageBox :parent => parent_window, :message => 'Отсутствует бумага в принтере', :title => "Ошибка подключения", :icon => 'error'
      return false
    elsif exellio.GetStatusBit(2,3) == 1
      exellio.CancelReceipt()
      if exellio.LastError > 0
        Tk::messageBox :parent => parent_window, :message => exellio.LastErrorText, :title => "Ошибка подключения", :icon => 'error'
        return false
      end
    elsif exellio.GetStatusBit(2,5) == 1
      exellio.CloseNonfiscalReceipt()
      if exellio.LastError > 0
        Tk::messageBox :parent => parent_window, :message => exellio.LastErrorText, :title => "Ошибка подключения", :icon => 'error'
        return false
      end
    end

    return (yield exellio)
  ensure
    exellio.ClosePort

    parent_window.cursor = ''
    parent_window.update
  end
rescue => ex
  Tk::messageBox :parent => parent_window, :message => "Ошибка создания объекта\nExellioFP.FiscalPrinter", :title => "Ошибка подключения", :icon => 'error', :detail => ex.message

  return false
end

def action(parent_window,name, sum=0)
  exellio_connect(parent_window) {|exellio|
    case name
      when "zero" then exellio.PrintNullCheck()
      when "last" then exellio.MakeReceiptCopy(1)
      when "x" then exellio.XReport("0000")
      when "z" then exellio.ZReportWC("0000")
      when "in" then exellio.InOut(sum)
      when "out" then exellio.InOut(-sum)
    end
  
    if exellio.LastError > 0
      Tk::messageBox :parent => parent_window, :message => exellio.LastErrorText, :title => "Ошибка печати", :icon => 'error'
      
      false
    else
      true
    end
  }
end

def plural(num,str1,str2,str3)
  ost = num - (num/10).to_i*10
  
  if ((num > 10)&(num < 20))||((num - (num/100).to_i*100 > 10)&(num - (num/100).to_i*100 < 20)) then str3
  elsif ost == 1 then str1
  elsif (ost > 1)&(ost < 5) then str2
  else str3
  end
end

def frm_sum(sum)
  ("%.2f"%sum).gsub(/(\d)(?=(\d{3})+(?!\d))/, "\\1 ")
end

def colorise_tree(parent)
  $tree.children(parent).each {|item|
    $tree.tag_remove('dir',item)
    $tree.tag_remove('errors',item)
    $tree.tag_remove('even',item)
    $tree.tag_remove('odd',item)
    
    if $tree.get(item,'errors').to_i > 0 then
      $tree.tag_add('errors',item)
    elsif $tree.children(item).length > 0
      $tree.tag_add('dir',item)
    else
      $tree.tag_add((item.index+1)%2 == 0 ? 'even':'odd',item)
    end
  }
  
  unless parent == $tree.root
    colorise_tree(parent.parent_item)
  end
end

def recount_info_variable()
  sum = 0
  docs_count = 0
  errors = 0

  $tree.root.children.each{|item|
    sum += item.get('sum').gsub(/\s+/, '').to_f
    docs_count += item.get('docs_count').to_i
    errors += item.get('errors').to_i
  }
  
  $total_sum_info.value = (sum != 0)?"Сумма: #{frm_sum(sum)}":''
  $total_documents_info.value = (docs_count>0)?"Документов: #{docs_count}":''
  $total_errors_info.value = (errors>0)?"Ошибки: #{errors}":''
end

def recount_parent_values(parent)
  sum = 0
  docs_count = 0
  errors = 0

  parent.children.each{|item|
    sum += item.get('sum').gsub(/\s+/, '').to_f
    docs_count += item.get('docs_count').to_i
    errors += item.get('errors').to_i
  }

  parent['text'] = "#{parent['text'][/.*[^\(\d*\/*\d*\)]/]}(#{docs_count}#{(errors>0)?("/"+errors.to_s):''})"
  parent.set('sum',frm_sum(sum))
  parent.set('docs_count',docs_count)
  parent.set('errors',errors)
    
  unless parent == $tree.root
    recount_parent_values(parent.parent_item)
  end
end

def move_file(filename,from_dir,to_dir)
  return unless File.exist?("#{from_dir}/#{filename}")
    
  dir, name = File.split(filename)
  
  path = "#{to_dir}";
  dir.split("\/").each {|d|
    path+="/#{d}"

    if not(File.exist?(path)) then
      Dir.mkdir(path)
    end
  }
 
  File.rename("#{from_dir}/#{filename}" ,"#{to_dir}/#{filename}")  
  
end

def menu_open_document_clicked()
  item = $tree.focus_item

  if not(item)
    Tk::messageBox :parent => $root, :message => 'Не выбран документ для печати чека!', :title => "Ошибка пользователя", :icon => 'info'

    return
  elsif not(File.file?("prepared/#{item.id}"))
    Tk::messageBox :parent => $root, :message => "#{item['text']} не является документом!", :title => "Ошибка пользователя", :icon => 'info'

    return
  end
  
  action_window = Tk::Toplevel.new($root) {withdraw; borderwidth 5; title item.id; resizable false,false; transient $root; iconphoto $root.iconphoto; grab}

  label = Tk::Tile::Label.new(action_window) {text "Документ:"; grid :column => 0, :row => 0, :sticky => 'w', :pady => 5}
  label = Tk::Tile::Label.new(action_window) {text item['text']; relief 'solid'; grid :column => 1, :row => 0, :sticky => 'we', :pady => 5}
  label = Tk::Tile::Label.new(action_window) {text "Сумма:"; grid :column => 2, :row => 0, :sticky => 'e', :padx => [5,0], :pady => 5}
  label = Tk::Tile::Label.new(action_window) {text item.get('sum'); relief 'solid'; anchor 'e'; grid :column => 3, :row => 0, :sticky => 'we', :pady => 5}
    
  tk_variable = TkVariable.new(-1)
  
  label = Tk::Tile::Label.new(action_window) {text "Сумма оплаты:"; grid :column => 0, :row => 1, :sticky => 'w'}
  entry = Tk::Tile::Entry.new(action_window) {width 12; focus; justify 'right'; validate 'key'; grid :column => 1, :row => 1, :columnspan => 3,:sticky => 'ew', :ipadx => 40}
  Tk::Tile::Checkbutton.new(entry) {text '...'; takefocus 0; variable tk_variable; command proc{calculator(action_window,entry,tk_variable,2)}; style 'Toolbutton'; cursor 'arrow'; pack :side => 'left', :padx => 1, :pady => 1}

  notebook_frame = Tk::Tile::Notebook.new(action_window) {grid :column => 0, :row => 2, :columnspan => 4, :pady => 5}

  frame_bills_counter, bills_count, coins_count, bills_entrys, coins_entrys = get_frame_bills_counter(notebook_frame,entry)
  frame_bills_counter.pack(:fill => 'both', :expand => 1)
  
  notebook_frame.add(frame_bills_counter,:text => 'Оплата')
  
  entry.validatecommand(proc {|p| 
    if check_entry(p,2)
      bills_count.each_key {|key| bills_count[key] = 0}
      coins_count.each_key {|key| coins_count[key] = 0}

      bills_entrys.each {|entr| entr.validate('none'); entr.value = ''; entr.validate('key')}
      coins_entrys.each {|entr| entr.validate('none'); entr.value = ''; entr.validate('key')}
        
      true
    else false
    end
    },'%P')
    
  frame_goods = Tk::Tile::Frame.new(notebook_frame) {pack :fill => 'both'}
  notebook_frame.add(frame_goods,:text => 'Товары')
  
  text = nil
  scroll_y = Tk::Tile::Scrollbar.new(frame_goods) {orient 'vertical'; command proc{|*args| text.yview(*args)}}
  scroll_x = Tk::Tile::Scrollbar.new(frame_goods) {orient 'horizontal'; command proc{|*args| text.xview(*args)}}
  text = Tk::Text.new(frame_goods) {yscrollcommand proc{|*args| scroll_y.set(*args)}; xscrollcommand proc{|*args| scroll_x.set(*args)}; width 30; height 18; wrap 'none'; tabs '1c 2c'}

  text.grid(:column => 0, :row => 0, :sticky => 'news')
  scroll_y.grid(:column => 1, :row => 0, :sticky => 'ns')
  scroll_x.grid(:column => 0, :row => 1, :sticky => 'ew')
  
  file = File.open("prepared/#{item.id}","r:cp1251")
  txt = file.readlines
  file.close
  
  txt.each_with_index do |str,index|
    items = str.split(/\s*;\s*/)

    sum = 0
    errors = 0
    if items.length == 6 then sum = items[4].to_f*items[5].to_f
    else errors = 1
    end

    text.insert('end',"#{index+1} (#{items[0]})\t#{items[1]}\n\t#{items[5]} * #{frm_sum(items[4])} = #{frm_sum(sum)}\n",((index+1).even?)?'even':'odd')
  end
  
  text.tag_configure('odd',:background => 'snow2')
  text['state'] = 'disabled'

  notebook_frame.itemconfigure(1,:text => "Товары (#{txt.length})")

  frame_buttons = Tk::Tile::Frame.new(action_window) {grid :column => 0, :row => 3, :columnspan => 4, :sticky => 'ew'}
  button_ok = Tk::Tile::Button.new(frame_buttons) {text "Печать"; command proc {
      sum = entry.value.sub(',','.').to_f
      if sum > 0
        if sum >= item.get('sum').delete("\s").to_f
          if print_document(action_window,"prepared/#{item.id}",sum)
            move_file(item.id,"prepared","printed")

            parent_item = item.parent_item

            $tree.delete(item)

            recount_parent_values(parent_item)

            colorise_tree(parent_item)

            while (parent_item.children.length == 0)&(parent_item != $tree.root)
              id = parent_item.id
              parent_item = parent_item.parent_item

              $tree.delete(id)

              if File.directory?("prepared/#{id}")
                Dir["prepared/#{id}/*"].each {|file| delete_file(file)}
                Dir.rmdir("prepared/#{id}")
              end
            end

            action_window.destroy
          end
        else
          Tk::messageBox :parent => action_window, :message => "Сумма оплаты (#{("%.2f"%sum)}) не может быть меньше суммы документа (#{item.get('sum')})",:title => "Ошибка печати чека", :icon => 'info'
        end
      else
        Tk::messageBox :parent => action_window, :message => "Сумма оплаты должна быть больше 0.00",:title => "Ошибка печати чека", :icon => 'info'
      end
    }
    grid :column => 1, :row => 0, :sticky => 'ew'}
  button_cancel = Tk::Tile::Button.new(frame_buttons) {text "Отмена"; command proc {action_window.destroy;  cancel = true}; grid :column => 2, :row => 0, :sticky => 'ew'}

  frame_buttons.grid_columnconfigure(0, :weight => 1)
  frame_buttons.grid_columnconfigure(1, :uniform => '1')
  frame_buttons.grid_columnconfigure(2, :uniform => '1')
  
  action_window.grid_columnconfigure(1, :weight => 1)
  action_window.grid_columnconfigure(3, :weight => 1)
  action_window.grid_columnconfigure(1, :uniform => '2')
  action_window.grid_columnconfigure(3, :uniform => '2')

  action_window.update

  action_window.geometry("+#{$root.winfo_x+($root.winfo_width-action_window.winfo_reqwidth)/2}+#{$root.winfo_y+($root.winfo_height-action_window.winfo_reqheight)/2}")

  action_window.deiconify

  action_window.wait_window
end

def menu_official_action_clicked(name)
  action_window = Tk::Toplevel.new($root) {withdraw; title (name=='in')?'Служебное внесение':'Служебная выдача'; resizable false,false; transient $root; iconphoto $root.iconphoto; grab}

  tk_variable = TkVariable.new(-1)
  
  label = Tk::Tile::Label.new(action_window) {text "Сумма #{(name=='in'?'внесения':'выдачи') }:"; grid :column => 0, :row => 1, :sticky => 'w', :padx => [5,0], :pady => 5}
  entry = Tk::Tile::Entry.new(action_window) {width 12; focus; justify 'right'; validate 'key'; grid :column => 1, :row => 1, :sticky => 'ew', :padx => 5, :pady => 5, :ipadx => 40}
  Tk::Tile::Checkbutton.new(entry) {text '...'; takefocus 0; variable tk_variable; command proc{calculator(action_window,entry,tk_variable,2)}; style 'Toolbutton'; cursor 'arrow'; pack :side => 'left', :padx => 1, :pady => 1}

  frame_bills_counter, bills_count, coins_count, bills_entrys, coins_entrys = get_frame_bills_counter(action_window,entry)
  frame_bills_counter.grid(:column => 0, :row => 2, :columnspan => 2)
  
  entry.validatecommand(proc {|p| 
    if check_entry(p,2)
      bills_count.each_key {|key| bills_count[key] = 0}
      coins_count.each_key {|key| coins_count[key] = 0}

      bills_entrys.each {|entr| entr.validate('none'); entr.value = ''; entr.validate('key')}
      coins_entrys.each {|entr| entr.validate('none'); entr.value = ''; entr.validate('key')}
        
      true
    else false
    end
    },'%P')

  frame_buttons = Tk::Tile::Frame.new(action_window) {grid :column => 0, :row => 3, :columnspan => 2, :sticky => 'ew'}
  button_ok = Tk::Tile::Button.new(frame_buttons) {text "Печать"; command proc {
      sum = entry.value.sub(',','.').to_f
      if sum > 0
        action_window.destroy if action(action_window,name,sum) 
      else
        Tk::messageBox :parent => action_window, :message => "Сумма #{(name=='in')?'внесения':'выдачи'} (#{("%12.2f"%sum).lstrip}) должна быть больше 0.00",:title => "Ошибка внесения суммы", :icon => 'info'
      end
    }
    grid :column => 1, :row => 0, :sticky => 'ew', :padx => [5,0], :pady => 5}
  button_cancel = Tk::Tile::Button.new(frame_buttons) {text "Отмена"; command proc {action_window.destroy;  cancel = true}; grid :column => 2, :row => 0, :sticky => 'ew', :padx => 5, :pady => 5}

  frame_buttons.grid_columnconfigure(0, :weight => 1)
  frame_buttons.grid_columnconfigure(1, :uniform => '1')
  frame_buttons.grid_columnconfigure(2, :uniform => '1')
  
  action_window.grid_columnconfigure(1, :weight => 1)

  action_window.update

  action_window.geometry("+#{$root.winfo_x+($root.winfo_width-action_window.winfo_reqwidth)/2}+#{$root.winfo_y+($root.winfo_height-action_window.winfo_reqheight)/2}")

  action_window.deiconify

  action_window.wait_window
end

def menu_last_clicked()
  return if (Tk::messageBox :parent => $root, :message => 'Выполнить печать копии последнего чека?', :title => "Печать копии чека", :icon => 'question', :type => 'yesno') == 'no'

  action($root,'last')
end

def menu_zero_clicked()
  return if (Tk::messageBox :parent => $root, :message => 'Выполнить печать нулевого чека?', :title => "Печать нулевого чека", :icon => 'question', :type => 'yesno') == 'no'

  action($root,'zero')
end

def menu_x_clicked()
  return if (Tk::messageBox :parent => $root, :message => 'Выполнить печать X-отчета?', :title => "Печать X-отчета", :icon => 'question', :type => 'yesno') == 'no'

  action($root,'x')
end

def menu_z_clicked()
  return if (Tk::messageBox :parent => $root, :message => 'Выполнить печать Z-отчета?', :title => "Печать Z-отчета", :icon => 'question', :type => 'yesno') == 'no'

  action($root,'z')
end

def menu_fill_tree_clicked()
  fill_tree
end

def load_dir(dir,parent,label,progress)
  Dir["#{dir}*.csv"].each {|filename|
    file = File.open(filename,"r:cp1251")
    text = file.readlines
    file.close
      
    item = $tree.insert(parent,'end',:id =>filename,:text =>File.basename(filename,'.csv'), :values => [0,1,0], :tag => 'docs')

    label['text'] = "Загрузка заданий (#{(progress['value']+1).to_i})"
      
    sum = 0
    text.each do |str|
      items = str.split(/\s*;\s*/)

      if items.length == 6 then sum += items[4].to_f*items[5].to_f
      else item.set('errors',1)
      end
    end

    item.set('sum',frm_sum(sum))
      
    progress.step
  }

  Dir["#{dir}*"].each{|filename|
    if File.directory?(filename)
      item = $tree.insert(parent,'end',:id => filename,:text =>"#{File.basename(filename)} (0)", :values => [0,0,0], :open => (parent==$tree.root)?true:false)
        
      load_dir("#{filename}/",item,label,progress)

      recount_parent_values(item)
    end
  }

  colorise_tree(parent)
end

def fill_tree()
  TkPack.forget($tree)

  $tree.delete($tree.children($tree.root))
    
  fill_window = Tk::Toplevel.new($root) {withdraw; focus; overrideredirect true; resizable false,false; grab 'global'}
  fill_frame = Tk::Tile::Frame.new(fill_window) {relief 'solid';pack}
  label = Tk::Tile::Label.new(fill_frame) {text 'Загрузка заданий'; pack :side => 'top', :anchor => 's', :expand => 1, :padx => 5, :pady => 5}
  progress = Tk::Tile::Progressbar.new(fill_frame) {orient 'horizontal'; mode 'indeterminate'; maximum 10; pack :side => 'top', :padx => 5, :pady => 5, :anchor => 'n', :fill => 'x',:expand => 1}

  fill_window.update

  fill_window.geometry("+#{$root.winfo_x+($root.winfo_width-fill_window.winfo_reqwidth)/2}+#{$root.winfo_y+($root.winfo_height-fill_window.winfo_reqheight)/2}")

  fill_window.deiconify
  
  Dir.chdir('prepared') {load_dir('',$tree.root,label,progress)}
  
  fill_window.destroy

  $tree.pack(:side => 'left', :fill => 'both', :expand => 1)
  
  recount_info_variable
end

$comport = ""

if File.exist?("ExellioFP.conf")
  file = File.open("ExellioFP.conf","r:cp1251")
  text = file.readlines
  file.close

  text.each {|str| $comport = str.sub("COMPORT=","").chomp! if str.start_with? "COMPORT="}
else
  File.open("ExellioFP.conf","w") {|config| config.puts("COMPORT="); config.puts("DIR=")}
end

TkOption.add '*tearOff', 0

$root = TkRoot.new {withdraw; title "Exellio FP (фискальный принтер) 2.1"; iconphoto TkPhotoImage.new(:file => 'dpp350.gif') if File.exist?('dpp350.gif'); state 'zoomed'}

menubar = TkMenu.new($root)
  
printer_menu = TkMenu.new(menubar)
printer_menu.add :command, :label => 'Печать нулевого чека', :command => proc{menu_zero_clicked}
printer_menu.add :command, :label => 'Печать копии последнего чека', :command => proc{menu_last_clicked}
printer_menu.add :separator
printer_menu.add :command, :label => 'Служебное внесение', :command => proc{menu_official_action_clicked('in')}
printer_menu.add :command, :label => 'Служебная выдача', :command => proc{menu_official_action_clicked('out')}
printer_menu.add :separator
printer_menu.add :command, :label => 'Обновить список', :command => proc{Thread.new {menu_fill_tree_clicked}}
printer_menu.add :separator
#printer_menu.add :command, :label => 'Состояние'
#printer_menu.add :command, :label => 'Настройка'
#printer_menu.add :separator
printer_menu.add :command, :label => 'Закрыть', :command => proc{$root.destroy}

reports_menu = TkMenu.new(menubar)
reports_menu.add :command, :label => 'X-отчет', :command => proc{menu_x_clicked}
reports_menu.add :command, :label => 'Z-отчет', :command => proc{menu_z_clicked}
#reports_menu.add :separator
#reports_menu.add :command, :label => 'Полный периодический отчет по номерам Z-отчетов'
#reports_menu.add :command, :label => 'Полный периодический отчет по датам Z-отчетов'
#reports_menu.add :separator
#reports_menu.add :command, :label => 'Краткий периодический отчет по номерам Z-отчетов'
#reports_menu.add :command, :label => 'Краткий периодический отчет по датам Z-отчетов'

menubar.add :cascade, :menu => printer_menu, :label => 'Принтер'
menubar.add :cascade, :menu => reports_menu, :label => 'Отчеты'

$root['menu'] = menubar

tree_frame = Tk::Tile::Frame.new($root)
  
$tree = nil
scroll = Tk::Tile::Scrollbar.new(tree_frame) {orient 'vertical'; command proc{|*args| $tree.yview(*args)}}
$tree = Tk::Tile::Treeview.new(tree_frame) {yscrollcommand proc{|*args| scroll.set(*args)}; selectmode 'browse'; columns 'sum docs_count errors'; displaycolumns 'sum'}

$tree.heading_configure('#0', :text => 'Документ')
$tree.heading_configure('sum', :text => 'Сумма')

$tree.column_configure('#0', :width => 200)
$tree.column_configure('sum',:anchor => 'e', :width => 90)

$tree.tag_configure('errors',:background => 'red')
$tree.tag_configure('odd',:background => 'snow2')
$tree.tag_configure('dir',:background => 'beige')

$tree.tag_bind('docs', 'Double-1', proc{menu_open_document_clicked});
$tree.tag_bind('docs', 'ButtonPress', proc{|x,y| check_for_open(x,y)},"%x %y");

#$tree.pack(:side => 'left', :fill => 'both', :expand => 1)
scroll.pack(:side => 'right', :fill => 'y')

statusbar_frame = Tk::Tile::Frame.new($root)

$total_sum_info = TkVariable.new
Tk::Tile::Label.new(statusbar_frame) {textvariable $total_sum_info; anchor 'w'; pack :side => 'left', :ipadx => 5}

$total_documents_info = TkVariable.new
Tk::Tile::Label.new(statusbar_frame) {textvariable $total_documents_info; anchor 'center'; pack :side => 'left', :ipadx => 5}

$total_errors_info = TkVariable.new
Tk::Tile::Label.new(statusbar_frame) {textvariable $total_errors_info; anchor 'center'; pack :side => 'left', :ipadx => 5}

tree_frame.grid(:column => 0, :row => 0, :sticky => 'news', :columnspan => 2)
statusbar_frame.grid(:column => 0, :row => 1, :sticky => 'ew')

Tk::Tile::SizeGrip.new($root) {grid :column => 1, :row => 1, :sticky => 'se'}

$root.grid_columnconfigure(0,:weight => 1)
$root.grid_rowconfigure(0,:weight => 1)

fill_tree

Tk.mainloop()