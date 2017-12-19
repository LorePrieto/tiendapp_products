require 'spec_helper'
require 'roo'

RSpec.describe TiendappProducts::Export do
  describe "A excel is created with the correct content from Spree" do
      it "should create an excel with correct content in simple case" do
        # We create an object in the database
        #Product
        sc = Spree::ShippingCategory.create!(name:"Por defecto")
        pr = Spree::Product.create!(name:"queque", price:1500, description: "en molde de cupcake", slug: "queque-en-molde-de-cupcake", meta_description: "Este es el mejor queque de Chile", available_on: "2017-12-06",  shipping_category_id: sc.id)
        #Properties
        p1 = Spree::Property.create!(name: "Hecho en casa", presentation: "Hecho en casa")
        p2 = Spree::Property.create!(name: "For real no fake", presentation: "For real no fake")
        Spree::ProductProperty.create!(value: "Sí", product_id: pr.id, property_id: p1.id)
        Spree::ProductProperty.create!(value: "No", product_id: pr.id, property_id: p2.id)
        #Options
        opt1 = Spree::OptionType.create!(name: "Sabor", presentation: "Sabor", position: 1)
        opt2 = Spree::OptionType.create!(name: "Tamaño", presentation: "Tamaño", position: 2)
        Spree::ProductOptionType.create!(position: 1, product_id: pr.id, option_type_id: opt1.id)
        Spree::ProductOptionType.create!(position: 2, product_id: pr.id, option_type_id: opt2.id)
        opv1 = Spree::OptionValue.create!(position: 1, name: "Nueces", presentation: "Nueces", option_type_id: opt1.id)
        opv2 = Spree::OptionValue.create!(position: 2, name: "Vainilla", presentation: "Vainilla", option_type_id: opt1.id)
        opv3 = Spree::OptionValue.create!(position: 1, name: "estoy a dieta", presentation: "estoy a dieta", option_type_id: opt2.id)
        opv4 = Spree::OptionValue.create!(position: 2, name: "gigante", presentation: "gigante", option_type_id: opt2.id)
        #Variants
        vr = Spree::Variant.create!(sku: "12345678", weight: 200, height: 10, width: 10, depth: 10, is_master: false, product_id: pr.id)
        vr.price = 2000
        vr.save!
        vr.option_values << opv1
        vr.option_values << opv3
        #Locations
        cr = Spree::Country.create!(iso_name: "CHILE", iso: "CL", iso3: "CHL", name:"Chile", numcode:"152", states_required:true)
        sta = Spree::State.create!(name: "Región Metropolitana", abbr: "RM", country_id: cr.id)
        loc = Spree::StockLocation.create!(name: "Isla Diamante", default: false, address1: "Playa 123", address2: "Juan algo 234", city: "Til Til", state_id: sta.id, country_id: cr.id, zipcode: "12345", phone:"76543469", active: true, backorderable_default: true, propagate_all_variants: true, admin_name: "Central")
        #Stock
        Spree::StockMovement.create!(stock_item_id: vr.stock_items.first.id, quantity: 20)

        # We ask the gem to create the excel
        TiendappProducts::Export.get_products('spec/fixtures/exported.xlsx')

        # We check that the excel holds the correct values
        #Headers
        xlsx = Roo::Spreadsheet.open('spec/fixtures/exported.xlsx')
        expect(xlsx.sheets).to eql(["Productos", "Propiedades", "Opciones", "Variantes", "Ubicaciones", "Stock"])
        expect(xlsx.sheet("Productos").row(1)).to eql(["ID*", "Nombre*", "Descripción", "Precio Principal*", "SKU Producto", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías",  "Categoría de Shipping*" ])
        expect(xlsx.sheet("Propiedades").row(1)).to eql(["ID Producto*", "Propiedad*", "Valor*"])
        expect(xlsx.sheet("Opciones").row(1)).to eql(["ID Producto*", "Opción*", "Valores*" ])
        expect(xlsx.sheet("Variantes").row(1)).to eql(["ID*", "ID Producto*", "Opciones*", "SKU Variante", "Precio*", "Peso", "Altura", "Longitud", "Profundidad"])
        expect(xlsx.sheet("Ubicaciones").row(1)).to eql(["Nombre*", "Nombre Interno*", "Calle*", "Ciudad*", "Calle de referencia", "Código Postal*", "Teléfono", "País*", "Región*", "Activa", "Por defecto", "Backorderable", "Propagar por todas las variantes"])
        expect(xlsx.sheet("Stock").row(1)).to eql(["ID Producto*", "Ubicación (Nom. Interno)", "ID Variante*", "Cantidad*", "Backorderable"])
        #Products
        expect(xlsx.sheet("Productos").row(2)).to eql([1, "queque", "en molde de cupcake", 1500.0, "N/A", "N/A", "N/A", "N/A", "N/A", "queque-en-molde-de-cupcake", "Este es el mejor queque de Chile", "Sí", "2017-12-06", "", "Por defecto" ])
        #Properties
        expect(xlsx.sheet("Propiedades").row(2)).to eql([1, "Hecho en casa", "Sí"])
        expect(xlsx.sheet("Propiedades").row(3)).to eql([1, "For real no fake", "No"])
        #Options
        expect(xlsx.sheet("Opciones").row(2)).to eql([1, "Sabor", "Nueces, Vainilla"])
        expect(xlsx.sheet("Opciones").row(3)).to eql([1, "Tamaño", "estoy a dieta, gigante"])
        #Variants
        expect(xlsx.sheet("Variantes").row(2)).to eql([2, 1, "Nueces, estoy a dieta", 12345678, 2000.0, 200.0, 10.0, 10.0, 10.0])
        #Locations
        expect(xlsx.sheet("Ubicaciones").row(2)).to eql(["Isla Diamante", "Central", "Playa 123", "Til Til", "Juan algo 234", 12345, 76543469, "Chile","Región Metropolitana", "Sí", "No", "Sí", "Sí"])
        #Stock
        expect(xlsx.sheet("Stock").row(2)).to eql([1, "Central", 2, 20, "Sí"])
      end
      it "should only add SKU, weight, etc in products tab if there is no variants " do
        # We create an object in the database
        #Product
        sc = Spree::ShippingCategory.create!(name:"Por defecto")
        pr = Spree::Product.create!(name:"queque", price:1500, description: "en molde de cupcake", slug: "queque-en-molde-de-cupcake", meta_description: "Este es el mejor queque de Chile", available_on: "2017-12-06",  shipping_category_id: sc.id)
        var = Spree::Variant.where(product_id: pr.id).first
        var.sku = "234566"
        var.width = 134
        var.height = 132
        var.weight = 123
        var.depth = 234
        var.save!

        # We ask the gem to create the excel
        TiendappProducts::Export.get_products('spec/fixtures/exported.xlsx')

        # We check that the excel holds the correct values
        xlsx = Roo::Spreadsheet.open('spec/fixtures/exported.xlsx')
        #Products
        expect(xlsx.sheet("Productos").row(2)).to eql([1, "queque", "en molde de cupcake", 1500.0, 234566, 123.0, 132.0, 134.0, 234.0, "queque-en-molde-de-cupcake", "Este es el mejor queque de Chile", "Sí", "2017-12-06", "", "Por defecto" ])
      end
      it "should print No in visible" do
        # We create an object in the database
        #Product
        t = (Date.current + 3).strftime("%Y-%m-%d")
        sc = Spree::ShippingCategory.create!(name:"Por defecto")
        pr = Spree::Product.create!(name:"queque", price:1500, description: "en molde de cupcake", slug: "queque-en-molde-de-cupcake", meta_description: "Este es el mejor queque de Chile", available_on: t,  shipping_category_id: sc.id)

        # We ask the gem to create the excel
        TiendappProducts::Export.get_products('spec/fixtures/exported.xlsx')

        # We check that the excel holds the correct values
        xlsx = Roo::Spreadsheet.open('spec/fixtures/exported.xlsx')
        #Products
        expect(xlsx.sheet("Productos").row(2)[11]).to eql("No")
      end
      it "Should print the categories (taxons)" do
        # We create an object in the database
        #Product
        sc = Spree::ShippingCategory.create!(name:"Por defecto")
        pr = Spree::Product.create!(name:"queque", price:1500, description: "en molde de cupcake", slug: "queque-en-molde-de-cupcake", meta_description: "Este es el mejor queque de Chile", available_on: "2017-12-06",  shipping_category_id: sc.id)
        #Categories
        ty = Spree::Taxonomy.create!(name: "riri")
        t1 = Spree::Taxon.where(taxonomy_id: ty.id).first
        t2 = Spree::Taxon.create(name:"ruru", parent_id: t1.id, taxonomy_id: ty.id)
        t3 = Spree::Taxon.create(name:"rere", parent_id: t2.id, taxonomy_id: ty.id)
        pr.taxons << t3

        # We ask the gem to create the excel
        TiendappProducts::Export.get_products('spec/fixtures/exported.xlsx')

        # We check that the excel holds the correct values
        xlsx = Roo::Spreadsheet.open('spec/fixtures/exported.xlsx')
        #Products
        expect(xlsx.sheet("Productos").row(2)[13]).to eql("riri->ruru->rere")
      end
      it "should fail if the file to import is not a xlsx" do
        m = TiendappProducts::Export.get_products('spec/fixtures/exported.xls')
        expect(m).to eql("Error: la extensión del archivo debe ser xlsx")
      end
  end
end
