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
          if row[4].to_s != "N/A" || row[5].to_s != "N/A" || row[6].to_s != "N/A" || row[7].to_s != "N/A" || row[8].to_s != "N/A"
            vr = Spree::Variant.where(product_id: pr.id, is_master: true).first
            if row[4].to_s != "N/A"
              vr.sku = row[4].to_i.to_s
            end
            if row[5].to_s != "N/A"
              vr.weight = row[5].to_f
            end
            if row[6].to_s != "N/A"
              vr.height = row[6].to_f
            end
            if row[7].to_s != "N/A"
              vr.width = row[7].to_f
            end
            if row[8].to_s != "N/A"
              vr.depth = row[8].to_f
            end
            vr.save
          end
        end
        #Properties
        ((xlsx.sheet("Propiedades").first_row + 1)..xlsx.sheet("Propiedades").last_row.to_i).each do |r|
          row = xlsx.sheet("Propiedades").row(r)
          pr = Spree::Product.where(slug: prod_dic[row[0].to_i]).first
          property = Spree::Property.create!(name: row[1].to_s, presentation: row[2].to_s)
          pr.properties << property
        end
        #Options
        opt_dic = Hash.new([])
        ((xlsx.sheet("Opciones").first_row + 1)..xlsx.sheet("Opciones").last_row.to_i).each do |r|
          row = xlsx.sheet("Opciones").row(r)
          pr = Spree::Product.where(slug: prod_dic[row[0].to_i]).first
          ot = Spree::OptionType.where(name: row[1].to_s).any? ? Spree::OptionType.where(name: row[1].to_s).first : Spree::OptionType.create!(name: row[1].to_s, presentation: row[1].to_s)
          Spree::ProductOptionType.create!(position: 1, product_id: pr.id, option_type_id: ot.id)
          values = row[2].split(',')
          values.each do |v|
            opv1 = Spree::OptionValue.create!(name: v.strip, presentation: v.strip, option_type_id: ot.id)
            opt_dic[row[0].to_i] << opv1
          end
        end
        #Variants
        var_dic = {}
        ((xlsx.sheet("Variantes").first_row + 1)..xlsx.sheet("Variantes").last_row.to_i).each do |r|
          row = xlsx.sheet("Variantes").row(r)
          pr = Spree::Product.where(slug: prod_dic[row[1].to_i]).first
          vr = Spree::Variant.create!(sku: row[3].to_i.to_s, weight: row[5].to_f, height: row[6].to_f, width: row[7].to_f, depth: row[8].to_f, is_master: false, product_id: pr.id)
          vr.price = row[4].to_f
          vr.save!
          values = row[2].split(',')
          op_vals = opt_dic[row[1].to_i]
          values.each do |v|
            op_vals.each do |opt|
              if opt.name == v.strip
                vr.option_values << opt
                break
              end
            end
          end
          var_dic[row[0].to_i] = vr.id
        end
        #Locations
        ((xlsx.sheet("Ubicaciones").first_row + 1)..xlsx.sheet("Ubicaciones").last_row.to_i).each do |r|
          row = xlsx.sheet("Ubicaciones").row(r)
          country = Spree::Country.where(name: row[7].to_s).any? ? Spree::Country.where(name: row[7].to_s).first : Spree::Country.create!(name: row[7].to_s, iso_name: row[7].to_s.upcase, states_required: true)
          sta = Spree::State.create!(name: row[8].to_s, country_id: country.id)
          loc = Spree::StockLocation.create!(name: row[0].to_s, admin_name: row[1].to_s, address1: row[2].to_s, city: row[3].to_s, address2: row[4].to_s, zipcode: row[5].to_i.to_s, phone: row[6].to_i.to_s,
             country_id: country.id, state_id: sta.id, active: row[9].to_s == "Sí", default: row[10].to_s == "Sí", backorderable_default: row[11].to_s == "Sí", propagate_all_variants: row[12].to_s == "Sí")
        end
        #Stock
        ((xlsx.sheet("Stock").first_row + 1)..xlsx.sheet("Stock").last_row.to_i).each do |r|
          row = xlsx.sheet("Stock").row(r)
          loc = Spree::StockLocation.where(admin_name: row[1].to_s).first
          vr = Spree::Variant.find(var_dic[row[2].to_i])
          item = Spree::StockItem.where(stock_location_id: loc.id, variant_id: vr.id).any? ? Spree::StockItem.where(stock_location_id: loc.id, variant_id: vr.id).first : Spree::StockItem.create!(stock_location_id: loc.id, variant_id: vr.id)
          item.backorderable = (row[4].to_s == "Sí")
          item.save
          Spree::StockMovement.create!(stock_item_id: item.id, quantity: row[3].to_i)
        end
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
      if xlsx.sheet("Ubicaciones").row(1) != ["Nombre", "Nombre Interno", "Calle", "Ciudad", "Calle de referencia", "Código Postal", "Teléfono", "País", "Región", "Activa", "Por defecto", "Backorderable", "Propagar por todas las variantes"]
        return "Los headers en la hoja Ubicaciones no son correctos"
      end
      if xlsx.sheet("Stock").row(1) != ["ID Producto", "Ubicación (Nom. Interno)", "ID Variante", "Cantidad", "Backorderable"]
        return "Los headers en la hoja Stock no son correctos"
      end
      return false
    end
  end
end
