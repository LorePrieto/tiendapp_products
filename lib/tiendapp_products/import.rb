module TiendappProducts
  class Import
    require 'roo'

    def self.create_products(path)
      xlsx = Roo::Spreadsheet.open(path)
      if !self.is_not_a_correct_excel(xlsx)
        # Products
        prod_dic = {}
        ((xlsx.sheet("Productos").first_row + 1)..xlsx.sheet("Productos").last_row.to_i).each do |r|
          row = xlsx.sheet("Productos").row(r)
          if row[14] != "Por defecto"
            sc = Spree::ShippingCategory.create!(name: row[14].to_s)
          else
            sc = Spree::ShippingCategory.where(name: "Por defecto").any? ? Spree::ShippingCategory.where(name: "Por defecto").first : Spree::ShippingCategory.create!(name: "Por defecto")
          end
          pr = Spree::Product.create!(name: row[1].to_s, description: row[2].to_s, price: row[3].to_i, slug: row[9].to_s, meta_description: row[10].to_s, available_on: DateTime.parse(row[12].to_s).to_date,  shipping_category_id: sc.id)
          prod_dic[row[0].to_i] = row[9].to_s
        end
        #Properties
        ((xlsx.sheet("Propiedades").first_row + 1)..xlsx.sheet("Propiedades").last_row.to_i).each do |r|
          row = xlsx.sheet("Propiedades").row(r)
          pr = Spree::Product.where(slug: prod_dic[row[0].to_i]).first
          property = Spree::Property.create!(name: row[1].to_s, presentation: row[2].to_s)
          pr.properties << property
        end
        #Options
        ((xlsx.sheet("Opciones").first_row + 1)..xlsx.sheet("Opciones").last_row.to_i).each do |r|
          row = xlsx.sheet("Opciones").row(r)
          pr = Spree::Product.where(slug: prod_dic[row[0].to_i]).first
          ot = Spree::OptionType.where(name: row[1].to_s).any? ? Spree::OptionType.where(name: row[1].to_s).first : Spree::OptionType.create!(name: row[1].to_s, presentation: row[1].to_s)
          Spree::ProductOptionType.create!(position: 1, product_id: pr.id, option_type_id: ot.id)
          values = row[2].split(',')
          values.each do |v|
            opv1 = Spree::OptionValue.create!(name: v.strip, presentation: v.strip, option_type_id: ot.id)
          end
        end
        #variants
        #Locations
        #Stock
      end
    end

    def self.is_not_a_correct_excel(xlsx)
      #Has the correct tabs and number of them
      if xlsx.sheets.count != 6
        return "Deben ser 6 hojas en el excel"
      end
      if !(xlsx.sheets.include? "Productos")
        return "No hay una hoja de Productos"
      end
      if !(xlsx.sheets.include? "Propiedades")
        return "No hay una hoja de Propiedades"
      end
      if !(xlsx.sheets.include? "Opciones")
        return "No hay una hoja de Opciones"
      end
      if !(xlsx.sheets.include? "Variantes")
        return "No hay una hoja de Variantes"
      end
      if !(xlsx.sheets.include? "Ubicaciones")
        return "No hay una hoja de Ubicaciones"
      end
      if !(xlsx.sheets.include? "Stock")
        return "No hay una hoja de Stock"
      end
      #Has the correct headers in the tabs
      if xlsx.sheet("Productos").row(1) != ["ID", "Nombre", "Descripción", "Precio Principal", "SKU", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías",  "Categoría de Shipping"]
        return "Los headers en la hoja Productos no son correctos"
      end
      if xlsx.sheet("Propiedades").row(1) != ["ID Producto", "Propiedad", "Valor"]
        return "Los headers en la hoja Propiedades no son correctos"
      end
      if xlsx.sheet("Opciones").row(1) != ["ID Producto", "Opción", "Valores" ]
        return "Los headers en la hoja Opciones no son correctos"
      end
      if xlsx.sheet("Variantes").row(1) != ["ID", "ID Producto", "Opciones", "SKU", "Precio", "Peso", "Altura", "Longitud", "Profundidad"]
        return "Los headers en la hoja Variantes no son correctos"
      end
      if xlsx.sheet("Ubicaciones").row(1) != ["ID", "Nombre", "Nombre Interno", "Calle", "Ciudad", "Calle de referencia", "Código Postal", "Teléfono", "País", "Región", "Activa", "Por defecto", "Backorderable", "Propagar por todas las variantes"]
        return "Los headers en la hoja Ubicaciones no son correctos"
      end
      if xlsx.sheet("Stock").row(1) != ["ID Producto", "ID Variante", "Cantidad", "ID Ubicación", "Backorderable"]
        return "Los headers en la hoja Stock no son correctos"
      end
      return false
    end
  end
end
