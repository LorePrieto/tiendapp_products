module TiendappProducts
  class Import
    require 'roo'

    def self.create_products(path)
      xlsx = Roo::Spreadsheet.open(path)
      if !self.is_not_a_correct_excel(xlsx)
        # Products
        prod_dic = self.products(xlsx)
        #Properties
        self.properties(xlsx, prod_dic)
        #Options
        opt_dic = self.options(xlsx, prod_dic)
        #Variants
        var_dic = self.variants(xlsx, prod_dic, opt_dic)
        #Locations
        self.locations(xlsx)
        #Stock
        self.stock(xlsx, var_dic)
      end
    end

    def self.is_not_a_correct_excel(xlsx)
      #Has the correct tabs and number of them
      if xlsx.sheets.count != 6
        return "\e[41mError de formato: Deben ser 6 hojas en el excel\e[0m"
      end
      if !(xlsx.sheets.include? "Productos")
        return "\e[41mError de formato: No hay una hoja de Productos\e[0m"
      end
      if !(xlsx.sheets.include? "Propiedades")
        return "\e[41mError de formato: No hay una hoja de Propiedades\e[0m"
      end
      if !(xlsx.sheets.include? "Opciones")
        return "\e[41mError de formato: No hay una hoja de Opciones\e[0m"
      end
      if !(xlsx.sheets.include? "Variantes")
        return "\e[41mError de formato: No hay una hoja de Variantes\e[0m"
      end
      if !(xlsx.sheets.include? "Ubicaciones")
        return "\e[41mError de formato: No hay una hoja de Ubicaciones\e[0m"
      end
      if !(xlsx.sheets.include? "Stock")
        return "\e[41mError de formato: No hay una hoja de Stock\e[0m"
      end
      #Has the correct headers in the tabs
      if xlsx.sheet("Productos").row(1) != ["ID*", "Nombre*", "Descripción", "Precio Principal*", "SKU Producto", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías",  "Categoría de Shipping*"]
        return "\e[41mError de formato: Los encabezados en la hoja Productos no son correctos\e[0m"
      end
      if xlsx.sheet("Propiedades").row(1) != ["ID Producto*", "Propiedad*", "Valor*"]
        return "\e[41mError de formato: Los encabezados en la hoja Propiedades no son correctos\e[0m"
      end
      if xlsx.sheet("Opciones").row(1) != ["ID Producto*", "Opción*", "Valores*" ]
        return "\e[41mError de formato: Los encabezados en la hoja Opciones no son correctos\e[0m"
      end
      if xlsx.sheet("Variantes").row(1) != ["ID*", "ID Producto*", "Opciones*", "SKU Variante", "Precio*", "Peso", "Altura", "Longitud", "Profundidad"]
        return "\e[41mError de formato: Los encabezados en la hoja Variantes no son correctos\e[0m"
      end
      if xlsx.sheet("Ubicaciones").row(1) != ["Nombre*", "Nombre Interno*", "Calle*", "Ciudad*", "Calle de referencia", "Código Postal*", "Teléfono", "País*", "Región*", "Activa", "Por defecto", "Backorderable", "Propagar por todas las variantes"]
        return "\e[41mError de formato: Los encabezados en la hoja Ubicaciones no son correctos\e[0m"
      end
      if xlsx.sheet("Stock").row(1) != ["ID Producto*", "Ubicación (Nom. Interno)", "ID Variante*", "Cantidad*", "Backorderable"]
        return "\e[41mError de formato: Los encabezados en la hoja Stock no son correctos\e[0m"
      end
      return false
    end

    def self.products(xlsx)
      prod_dic = {}
      ((xlsx.sheet("Productos").first_row + 1)..xlsx.sheet("Productos").last_row.to_i).each do |r|
        row = xlsx.sheet("Productos").row(r)
        sc = Spree::ShippingCategory.where(name: row[14].to_s).any? ? Spree::ShippingCategory.where(name: row[14].to_s).first : Spree::ShippingCategory.create!(name: row[14].to_s)
        if row[9].to_s == ""
          pr = Spree::Product.where(slug: row[1].to_s.downcase .gsub(" ", "-")).any? ? Spree::Product.where(slug: row[1].to_s.downcase .gsub(" ", "-")).first : Spree::Product.create!(name: row[1].to_s, price: row[3].to_i, shipping_category_id: sc.id, slug: row[1].to_s.downcase .gsub(" ", "-"))
          prod_dic[row[0].to_i] = row[1].to_s.downcase .gsub(" ", "-")
        else
          pr = Spree::Product.where(slug: row[9].to_s).any? ? Spree::Product.where(slug: row[9].to_s).first : Spree::Product.create!(name: row[1].to_s, price: row[3].to_i, shipping_category_id: sc.id, slug: row[9].to_s)
          prod_dic[row[0].to_i] = row[9].to_s
        end
        pr.description = row[2].to_s
        pr.meta_description = row[10].to_s
        pr.available_on = row[12].to_s == "" ? nil : DateTime.parse(row[12].to_s).to_date
        pr.save
        if (row[4].to_s != "N/A" && row[4].to_s != "") ||
           (row[5].to_s != "N/A" && row[5].to_s != "") ||
           (row[6].to_s != "N/A" && row[6].to_s != "") ||
           (row[7].to_s != "N/A" && row[7].to_s != "") ||
           (row[8].to_s != "N/A" && row[8].to_s != "")
          vr = Spree::Variant.where(product_id: pr.id, is_master: true).first
          if (row[4].to_s != "N/A" && row[4].to_s != "")
            vr.sku = row[4].to_s
          end
          if (row[5].to_s != "N/A" && row[5].to_s != "")
            vr.weight = row[5].to_f
          end
          if (row[6].to_s != "N/A" && row[6].to_s != "")
            vr.height = row[6].to_f
          end
          if (row[7].to_s != "N/A" && row[7].to_s != "")
            vr.width = row[7].to_f
          end
          if (row[8].to_s != "N/A" && row[8].to_s != "")
            vr.depth = row[8].to_f
          end
          vr.save
        end
        taxons = row[13].to_s.split(', ')
        taxons.each do |taxon|
          taxs = taxon.split('->')
          ty = Spree::Taxonomy.where(name: taxs.first).any? ? Spree::Taxonomy.where(name: taxs.first).first : Spree::Taxonomy.create!(name: taxs.first)
          tp = Spree::Taxon.where(taxonomy_id: ty.id).first
          taxs.drop(1).each do |tax|
            tp = Spree::Taxon.where(name: tax, parent_id: tp.id, taxonomy_id: ty.id).any? ? Spree::Taxon.where(name: tax, parent_id: tp.id, taxonomy_id: ty.id).first : Spree::Taxon.create!(name: tax, parent_id: tp.id, taxonomy_id: ty.id)
          end
          if !pr.taxons.where(name: tp.name, parent_id: tp.parent_id, taxonomy_id: ty.id).any?
            pr.taxons << tp
          end
        end
      end
      return prod_dic
    end

    def self.properties(xlsx, prod_dic)
      ((xlsx.sheet("Propiedades").first_row + 1)..xlsx.sheet("Propiedades").last_row.to_i).each do |r|
        row = xlsx.sheet("Propiedades").row(r)
        pr = Spree::Product.where(slug: prod_dic[row[0].to_i]).first
        property = Spree::Property.where(name: row[1].to_s, presentation: row[1].to_s).any? ? Spree::Property.where(name: row[1].to_s, presentation: row[1].to_s).first : Spree::Property.create!(name: row[1].to_s, presentation: row[1].to_s)
        Spree::ProductProperty.where(value: row[2].to_s, product_id: pr.id, property_id: property.id).any? ? Spree::ProductProperty.where(value: row[2].to_s, product_id: pr.id, property_id: property.id).first : Spree::ProductProperty.create!(value: row[2].to_s, product_id: pr.id, property_id: property.id)
      end
    end

    def self.options(xlsx, prod_dic)
      opt_dic = Hash.new([])
      ((xlsx.sheet("Opciones").first_row + 1)..xlsx.sheet("Opciones").last_row.to_i).each do |r|
        row = xlsx.sheet("Opciones").row(r)
        pr = Spree::Product.where(slug: prod_dic[row[0].to_i]).first
        ot = Spree::OptionType.where(name: row[1].to_s).any? ? Spree::OptionType.where(name: row[1].to_s).first : Spree::OptionType.create!(name: row[1].to_s, presentation: row[1].to_s)
        Spree::ProductOptionType.where(product_id: pr.id, option_type_id: ot.id).any? ? Spree::ProductOptionType.where(product_id: pr.id, option_type_id: ot.id).first : Spree::ProductOptionType.create!(product_id: pr.id, option_type_id: ot.id)
        values = row[2].split(',')
        values.each do |v|
          opv1 = Spree::OptionValue.where(name: v.strip, presentation: v.strip, option_type_id: ot.id).any? ? Spree::OptionValue.where(name: v.strip, presentation: v.strip, option_type_id: ot.id).first : Spree::OptionValue.create!(name: v.strip, presentation: v.strip, option_type_id: ot.id)
          opt_dic[row[0].to_i] << opv1
        end
      end
      return opt_dic
    end

    def self.variants(xlsx, prod_dic, opt_dic)
      var_dic = {}
      ((xlsx.sheet("Variantes").first_row + 1)..xlsx.sheet("Variantes").last_row.to_i).each do |r|
        row = xlsx.sheet("Variantes").row(r)
        pr = Spree::Product.where(slug: prod_dic[row[1].to_i]).first
        vr = Spree::Variant.where(is_master: false, product_id: pr.id).any? ? Spree::Variant.where(is_master: false, product_id: pr.id).first : Spree::Variant.create!(is_master: false, product_id: pr.id)
        if (row[3].to_s != "N/A" && row[3].to_s != "") ||
           (row[5].to_s != "N/A" && row[5].to_s != "") ||
           (row[6].to_s != "N/A" && row[6].to_s != "") ||
           (row[7].to_s != "N/A" && row[7].to_s != "") ||
           (row[8].to_s != "N/A" && row[8].to_s != "")
          vr = Spree::Variant.where(product_id: pr.id, is_master: false).first
          if (row[3].to_s != "N/A" && row[4].to_s != "")
            vr.sku = row[3].to_s
          end
          if (row[5].to_s != "N/A" && row[5].to_s != "")
            vr.weight = row[5].to_f
          end
          if (row[6].to_s != "N/A" && row[6].to_s != "")
            vr.height = row[6].to_f
          end
          if (row[7].to_s != "N/A" && row[7].to_s != "")
            vr.width = row[7].to_f
          end
          if (row[8].to_s != "N/A" && row[8].to_s != "")
            vr.depth = row[8].to_f
          end
        end
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
      return var_dic
    end

    def self.locations(xlsx)
      ((xlsx.sheet("Ubicaciones").first_row + 1)..xlsx.sheet("Ubicaciones").last_row.to_i).each do |r|
        row = xlsx.sheet("Ubicaciones").row(r)
        country = Spree::Country.where(name: row[7].to_s).any? ? Spree::Country.where(name: row[7].to_s).first : Spree::Country.create!(name: row[7].to_s, iso_name: row[7].to_s.upcase, states_required: true)
        sta = Spree::State.where(name: row[8].to_s, country_id: country.id).any? ? Spree::State.where(name: row[8].to_s, country_id: country.id).first : Spree::State.create!(name: row[8].to_s, country_id: country.id)
        loc = Spree::StockLocation.where(admin_name: row[1].to_s).any? ? Spree::StockLocation.where(admin_name: row[1].to_s).first : Spree::StockLocation.create!(name: row[0].to_s, admin_name: row[1].to_s)
        loc.address1 = row[2].to_s
        loc.city = row[3].to_s
        loc.address2 = row[4].to_s
        loc.zipcode = row[5].to_i.to_s
        loc.phone = row[6].to_i.to_s == "" ? nil : row[6].to_s
        loc.country_id = country.id
        loc.state_id = sta.id
        loc.active = row[9].to_s == "" ? true : row[9].to_s == "Sí"
        loc.default = row[10].to_s == "" ? false : row[10].to_s == "Sí"
        loc.backorderable_default = row[11].to_s == "" ? false : row[11].to_s == "Sí"
        loc.propagate_all_variants = row[12].to_s == "" ? true : row[12].to_s == "Sí"
        loc.save
      end
    end

    def self.stock(xlsx, var_dic)
      ((xlsx.sheet("Stock").first_row + 1)..xlsx.sheet("Stock").last_row.to_i).each do |r|
        row = xlsx.sheet("Stock").row(r)
        loc = row[1].to_s == "" ? Spree::StockLocation.first : Spree::StockLocation.where(admin_name: row[1].to_s).first
        vr = Spree::Variant.find(var_dic[row[2].to_i])
        item = Spree::StockItem.where(stock_location_id: loc.id, variant_id: vr.id).any? ? Spree::StockItem.where(stock_location_id: loc.id, variant_id: vr.id).first : Spree::StockItem.create!(stock_location_id: loc.id, variant_id: vr.id)
        item.backorderable = row[4].to_s == "" ? false : row[4].to_s == "Sí"
        item.save
        if item.count_on_hand != row[3].to_i
          Spree::StockMovement.create!(stock_item_id: item.id, quantity: (row[3].to_i - item.count_on_hand))
        end
      end
    end

  end
end
