class ExcelsController < ActionController::Base
  def index; end

  def create
    excel_file_path = Rails.root.join('public', 'excel_reports', "Report---#{Time.now.strftime('%Y-%m-%d_%T')}.xlsx")

    Axlsx::Package.new do |p|
      p.workbook.add_worksheet(:name => "Work Hours") do |sheet|

        # Это делает формат со стороны гема
        date_style = p.workbook.styles.add_style(types: [:date], format_code: 'DD/MM/YY H:MM;@', alignment: { horizontal: :center })
        sheet.add_row [define_date], style: date_style
      end

      p.serialize(excel_file_path)
    end

    send_file excel_file_path
  end

  private

  def define_date
    # Поиграй с этой строчкой
    "23.12.2024 12:15".to_datetime.strftime("%d.%m.%y %H:%M")
  end
end
