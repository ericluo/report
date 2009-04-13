Shoes.setup do
  gem 'spreadsheet'
end

require 'spreadsheet'
require 'report_aggreator'

Shoes.app :title => "报表汇总" do
  stack :margin => 10 do
    para "请输入报表模板文件路径及文件名"
    flow do
      @template = edit_line "/home/eric/workspace/report/template.xls", :width => 400
      button "浏览" do
        @template.text = ask_open_file 
      end
    end

    para "请输入报表模板文件数据项范围"
    @range = edit_line "B7:F11", :width => 400

    para "请输入金融机构上报报表目录"
    flow do
      @reports_path = edit_line "/home/eric/workspace/report/data/", :width => 400
      button "浏览" do
        @reports_path.text = ask_open_folder
      end
    end

    para "请输入生成汇总文件路径及文件名"
    flow do
      @collect = edit_line "/home/eric/workspace/report/collect.xls", :width => 400
      button "浏览" do
        @collect.text = ask_open_file
      end
    end

    button "汇总" do
      aggreator = ReportAggreator.new(@template.text, @reports_path.text, 
                                      @collect.text, @range.text)
      @p = progress :width => 1.0
      @info = para "开始报表汇总"

      aggreator.aggreate do |source, index|
        debug("processing #{source}")
        @p.fraction = (index + 1) / aggreator.size
        if index == aggreator.size - 1
          @info.replace "汇总完成！"
        else
          @info.replace "汇总#{source}..."
        end
      end
    end
  end
end

#vim:ts=2 shiftwidth=2 softtabstop=2 expandtab:
