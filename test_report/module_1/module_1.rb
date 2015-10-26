require 'rubygems'
require 'roo'

flag=Array.new
colum=Array.new
type=Array.new
reg_exp=Array.new


file=File.new('C:\Users\nakul.therap\Documents\Aptana Studio 3 Workspace\lTest\test_data\module_1\expected_data', 'r' )

array_expected_data=Array.new

while line=file.gets
  if line[0]=="#"
  next
  else
    line_split=line.split(/[\s,=]/)

    line_split.each do |ls|
      array_expected_data<<ls
    end
  end
end
file.close

array_length=array_expected_data.length
#p array_length , array_expected_data

ods = Roo::Spreadsheet.open('C:\Users\nakul.therap\Documents\Aptana Studio 3 Workspace\lTest\test_data\module_1\spreadsheet1.ods')

ods_limit=0

ods.each do |ods_array|
  ods_limit+=1
end

file_out=File.new('C:\Users\nakul.therap\Documents\Aptana Studio 3 Workspace\lTest\test_report\module_1\output.html','w')

file_out.puts "<html><title>error report</title></head><body>"

for m in 0...array_length
  word=array_expected_data[m]
  limit=word.length-1

  if word[0...limit]=="column"
    colum = word[limit..limit].to_i
    type=array_expected_data[m+1]
    flag=array_expected_data[m+2]
    reg_exp=array_expected_data[m+3] 
    m+=3

    if colum==1
      for n in 2..ods_limit
        if flag == "true"
          s=ods.cell(n,colum)
          if s==nil
            file_out.puts "error in cell no <font color='red'> #{n}-#{colum}</font><br>"
          else
            s_limit=s.length
            for k in 0...s_limit
              if !((s[k]>='a' && s[k]<='z')||(s[k]>='A' && s[k]<='Z'))
                file_out.puts "error in cell no <font color='red'> #{n}-#{colum}</font><br>"
                break
              end 
            end
          end
        end
      end
    end
    
    if colum== 2
      for n in 2...ods_limit
        if flag== "true"
          s=ods.cell(n,colum)
          if s==nil
            file_out.puts "error in cell no <font color='red'> #{n}-#{colum}</font><br>"
          else
            s_limit=s.length
            if !(s[0...2].to_i>=1 && s[0...2].to_i<=31 && s[2...3]=='-' && s[3...5].to_i>=1 && s[3...5].to_i<=12 && s[5...6]=='-' && s[6...s_limit].to_i>=0 && s[6...s_limit].to_i<=9999)
              file_out.puts "error in cell no <font color='red'> #{n}-#{colum}</font><br>"
            end 
          end
        end
      end
    end
    
    if colum==3
      for n in 2..ods_limit
        if flag=="true"
          s=ods.cell(n,colum)
          if s==nil
            file_out.puts "error in cell no <font color='red'> #{n}-#{colum}</font><br>"
          else
            s_limit=s.length
              for k in 0...s_limit
                if !(s[k]>='0' && s[k]<='9')
                  file_out.puts "error in cell no <font color='red'> #{n}-#{colum}</font><br>"
                  break
                end 
              end
          end
        end
      end
    end
    file_out.puts "<br>"
    
  end
end

file_out.puts "</body>"
