module TiendappProducts
  class Export
    require 'axlsx'

    def self.get_products
      Axlsx::Package.new do |p|
        p.workbook.add_worksheet(:name => "Productos") do |sheet|
          sheet.add_row ["ID", "Nombre", "Descripción", "Precio Principal", "SKU", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías" ]
        end
        p.workbook.add_worksheet(:name => "Propiedades") do |sheet|
          sheet.add_row ["ID Producto", "Propiedad", "Valor"]
        end
        p.workbook.add_worksheet(:name => "Opciones") do |sheet|
          sheet.add_row ["ID Producto", "Opción", "Valores" ]
        end
        p.workbook.add_worksheet(:name => "Variantes") do |sheet|
          sheet.add_row ["ID", "ID Producto", "Opciones", "SKU", "Precio", "Peso", "Altura", "Longitud", "Profundidad"]
        end
        p.workbook.add_worksheet(:name => "Stock") do |sheet|
          sheet.add_row ["ID Producto", "ID Variante", "Cantidad", "Ubicación", "Backorderable"]
        end
        p.serialize('spec/fixtures/exported.xlsx')
      end
    end
  end
end
