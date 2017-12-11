module TiendappProducts
  class Export
    require 'axlsx'

##<Spree::Product id: nil, name: "Fran", description: nil, available_on: nil, deleted_at: nil, slug: nil, meta_description: nil, meta_keywords: nil, tax_category_id: nil, shipping_category_id: nil, created_at: nil, updated_at: nil, promotionable: true, meta_title: nil>


    def self.get_products
      Axlsx::Package.new do |p|
        p.workbook.add_worksheet(:name => "Productos") do |sheet|
          sheet.add_row ["ID", "Nombre", "Descripción", "Precio Principal", "SKU", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías" ]
          Spree::Product.all.each do |product|
            sheet.add_row [product.id, product.name, product.description, product.price, "TODO", "TODO", "TODO", "TODO", "TODO", product.slug, product.meta_description, "TODO", product.available_on.to_s, "TODO" ]
          end
        end
        p.workbook.add_worksheet(:name => "Propiedades") do |sheet|
          sheet.add_row ["ID Producto", "Propiedad", "Valor"]
          Spree::Product.all.each do |product|
            product.properties.each do |property|
              sheet.add_row [product.id, property.name, property.presentation ]
            end
          end
        end
        p.workbook.add_worksheet(:name => "Opciones") do |sheet|
          sheet.add_row ["ID Producto", "Opción", "Valores" ]
          Spree::Product.all.each do |product|
            product.option_types.each do |opt|
              values = opt.option_values
              val = values.first.name
              values.drop(1).each do |v|
                val = val + ", " + v.name
              end
              sheet.add_row [product.id, opt.name, val]
            end
          end
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
