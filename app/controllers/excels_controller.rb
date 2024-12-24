class ExcelsController < ActionController::Base
  def index; end

  def create
    excel_file_path = Rails.root.join('public', 'excel_reports', "Report---#{Time.now.strftime('%Y-%m-%d_%T')}.xlsx")

    Axlsx::Package.new do |p|
      p.workbook.add_worksheet(:name => "Work Hours") do |sheet|

        # Это делает формат со стороны гема
        date_style = p.workbook.styles.add_style(types: [:date], format_code: 'dd.mm.yy hh:mm', alignment: { horizontal: :center })

        # Если хочешь подключить формат со стороны гема, раскомментируй style: date_style
        #
        sheet.add_row [define_date], style: date_style
      end

      p.serialize(excel_file_path)
    end

    send_file excel_file_path
  end

  private

  def define_date
    # Поиграй с этой строчкой
    "23.12.2024 12:15:30".to_datetime.strftime("%d.%m.%y %H:%M")

    # Эти методы я нарыл и их пока изучаю, может быть с ними все сработает
  	# serial_date = Axlsx::DateTimeConverter.date_to_serial(DateTime.current)
  	# serial_date
  end
end
