require 'spec_helper'
require 'roo'
require 'tiendapp_products/export'

RSpec.describe Spree::Product do
  describe "Exported excel have correct headers" do
      it "should create an excel with correct headers" do
        # We create an object in the database
        #Product
        sc = Spree::ShippingCategory.create!(name:"Por defecto")
        pr = Spree::Product.create!(name:"queque", price:"1500", description: "en molde de cupcake", slug: "queque-en-molde-de-cupcake", meta_description: "Este es el mejor queque de Chile", available_on: "2017-12-06 17:26:02 UTC",  shipping_category_id: sc.id)
        #Properties
        p1 = Spree::Property.create!(name: "Hecho en casa", presentation: "Sí")
        p2 = Spree::Property.create!(name: "For real no fake", presentation: "No")
        pr.properties << p1
        pr.properties << p2
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
        #Spree::Price.create!(variant_id: vr.id, amount: 2000.0)
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
        TiendappProducts::Export.get_products()

        # We check that the excel holds the correct values
        #Headers
        xlsx = Roo::Spreadsheet.open('spec/fixtures/exported.xlsx')
        expect(xlsx.sheets).to eql(["Productos", "Propiedades", "Opciones", "Variantes", "Ubicaciones", "Stock"])
        expect(xlsx.sheet("Productos").row(1)).to eql(["ID", "Nombre", "Descripción", "Precio Principal", "SKU", "Peso", "Altura", "Longitud", "Profundidad", "Slug", "Descripción Meta", "Visible", "Disponible en", "Categorías" ])
        expect(xlsx.sheet("Propiedades").row(1)).to eql(["ID Producto", "Propiedad", "Valor"])
        expect(xlsx.sheet("Opciones").row(1)).to eql(["ID Producto", "Opción", "Valores" ])
        expect(xlsx.sheet("Variantes").row(1)).to eql(["ID", "ID Producto", "Opciones", "SKU", "Precio", "Peso", "Altura", "Longitud", "Profundidad"])
        expect(xlsx.sheet("Stock").row(1)).to eql(["ID Producto", "ID Variante", "Cantidad", "ID Ubicación", "Backorderable"])
        #Products
        expect(xlsx.sheet("Productos").row(2)).to eql([1.0, "queque", "en molde de cupcake", 1500.0, "TODO", "TODO", "TODO", "TODO", "TODO", "queque-en-molde-de-cupcake", "Este es el mejor queque de Chile", "TODO", "2017-12-06 17:26:02 UTC", "TODO" ])
        #Properties
        expect(xlsx.sheet("Propiedades").row(2)).to eql([1.0, "Hecho en casa", "Sí"])
        expect(xlsx.sheet("Propiedades").row(3)).to eql([1.0, "For real no fake", "No"])
        #Options
        expect(xlsx.sheet("Opciones").row(2)).to eql([1.0, "Sabor", "Nueces, Vainilla"])
        expect(xlsx.sheet("Opciones").row(3)).to eql([1.0, "Tamaño", "estoy a dieta, gigante"])
        #Variants
        expect(xlsx.sheet("Variantes").row(2)).to eql([2.0, 1.0, "Nueces, estoy a dieta", 12345678.0, 2000.0, 200.0, 10.0, 10.0, 10.0])
        #Locations
        expect(xlsx.sheet("Ubicaciones").row(2)).to eql([1.0, "Isla Diamante", "Central", "Playa 123", "Til Til", "Juan algo 234", 12345.0, 76543469.0, "Chile","Región Metropolitana", "Sí", "No", "Sí", "Sí"])
        #Stock
        expect(xlsx.sheet("Stock").row(2)).to eql([1.0, 2.0, 20.0, 1.0, "Sí"])
      end
  end
end
