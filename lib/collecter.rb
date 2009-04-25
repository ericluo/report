Shoes.setup do
  gem 'spreadsheet'
end

require 'spreadsheet'
require 'aggreator'

Shoes.app :title => "报表汇总" do

  [Shoes::Para, Shoes::Button, Shoes::Title].each do |element|
    style element, :font => "宋体" # "文泉驿正黑"
  end

  title "湖北银监局报表汇总生成器", :align => "center"
  stack :margin => 10 do
    para "请输入报表模板文件路径及文件名"
    flow do
      @template = edit_line "/home/eric/workspace/report/data/template.xls", :width => 400
      button "浏览" do
        @template.text = ask_open_file 
      end
    end

    para "请输入报表模板文件数据项范围"
    @range = edit_line "B7:F11", :width => 400

    para "请输入金融机构上报报表目录"
    flow do
      @reports_path = edit_line "/home/eric/workspace/report/data/report/", :width => 400
      button "浏览" do
        @reports_path.text = ask_open_folder
      end
    end

    para "请输入生成汇总文件路径及文件名"
    flow do
      @collect = edit_line "/home/eric/workspace/report/data/collect.xls", :width => 400
      button "浏览" do
        @collect.text = ask_open_file
      end
    end

    flow :margin_left => 320, :margin_top => 20 do
      clear_button = button "清除" 
      aggreator_button = button "汇总"
      button "退出" do exit; end

      clear_button.click {|btn, left, top|
        [@template, @range, @reports_path, @collect].each do |box|
          box.text = ""
        end
        @aggreator = nil
        @status_bar.clear
        aggreator_button.state = nil
      }

      aggreator_button.click {|btn, left, top|
        aggreator_button.state = "disabled"
        @status_bar = flow do
          @p = progress :width => 1.0
          @info = para ""
          updater = animate(24) do |frame|
            if @started
              current_index = @aggreator.processing_index
              @p.fraction = (current_index + 1) / @aggreator.size
              if current_index == @aggreator.size - 1
                @info.replace "汇总完成！"
                updater.stop
              else
                @info.replace "汇总#{source}..."
              end
            end
          end
        end
        @aggreator = Aggreator.new(@template.text, @reports_path.text, 
                                   @collect.text, @range.text)
        @started = true
        @aggreator.aggreate
      }
    end
  end
end

#vim:ts=2 shiftwidth=2 softtabstop=2 expandtab:
