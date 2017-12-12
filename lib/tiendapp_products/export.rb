module TiendappProducts
  class Export
    require 'axlsx'

    def self.get_products(path)
      Axlsx::Package.new do |p|
        p.workbook.add_worksheet(:name => "Productos") do |sheet|
          sheet.add_row ["ID", "Nombre", "Descripción", "Precio Principal", "SKU", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías", "Categoría de Shipping" ]
          Spree::Product.all.each do |product|
            var = Spree::Variant.where(product_id: product.id).where(is_master: true).first
            if product.variants.count > 0
              sheet.add_row [product.id, product.name, product.description, product.price, "N/A", "N/A", "N/A", "N/A", "N/A", product.slug, product.meta_description, product.available? ? "Sí" : "No", product.available_on.strftime("%Y-%m-%d"), "TODO", product.shipping_category.name]
            else
              sheet.add_row [product.id, product.name, product.description, product.price, var.sku, var.weight, var.height, var.width, var.depth, product.slug, product.meta_description, product.available? ? "Sí" : "No", product.available_on.strftime("%Y-%m-%d"), "TODO", product.shipping_category.name]
            end
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
          Spree::Product.all.each do |product|
            product.variants.each do |variant|
              values = variant.option_values
              val = values.first.name
              values.drop(1).each do |v|
                val = val + ", " + v.name
              end
              sheet.add_row [variant.id, product.id, val, variant.sku, variant.price, variant.weight, variant.height, variant.width, variant.depth]
            end
          end
        end
        p.workbook.add_worksheet(:name => "Ubicaciones") do |sheet|
          sheet.add_row ["ID", "Nombre", "Nombre Interno", "Calle", "Ciudad", "Calle de referencia", "Código Postal", "Teléfono", "País", "Región", "Activa", "Por defecto", "Backorderable", "Propagar por todas las variantes"]
          Spree::StockLocation.all.each do |location|
              sheet.add_row [location.id, location.name, location.admin_name, location.address1, location.city, location.address2, location.zipcode, location.phone, location.country.name, location.state.name, location.active ? "Sí" : "No", location.default ? "Sí" : "No", location.backorderable_default ? "Sí" : "No", location.propagate_all_variants ? "Sí" : "No" ]
          end
        end
        p.workbook.add_worksheet(:name => "Stock") do |sheet|
          sheet.add_row ["ID Producto", "ID Variante", "Cantidad", "ID Ubicación", "Backorderable"]
          Spree::Product.all.each do |product|
            product.variants.each do |variant|
              variant.stock_items.each do |stock|
                sheet.add_row [product.id, variant.id, stock.count_on_hand, stock.stock_location_id, stock.backorderable ? "Sí" : "No" ]
              end
            end
          end
        end
        p.serialize(path)
      end
    end
  end
end
