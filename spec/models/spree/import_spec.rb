require 'spec_helper'
require 'roo'

RSpec.describe TiendappProducts::Import do
  describe "From an excel the correct database enttries are created with Spree" do
    it "should create the correct database entries for simple case" do
      #We load the excel
      TiendappProducts::Import.create_products('spec/fixtures/imported.xlsx')

      #Check the database for the correct entries
      product = Spree::Product.where(slug: "queque-en-molde-de-cupcake").first
      sc = Spree::ShippingCategory.where(name: "Por defecto").first
      #Products
      expect(product.name).to eql("queque")
      expect(product.description).to eql("en molde de cupcake")
      expect(product.price).to eql(1500)
      expect(product.meta_description).to eql("Este es el mejor queque de Chile")
      expect(product.available_on.strftime("%Y-%m-%d")).to eql(DateTime.parse("2017-12-06 17:26:02 UTC").strftime("%Y-%m-%d"))
      expect(product.shipping_category_id).to eql(sc.id)
      #Propiedades
      expect(product.properties.first.name).to eql("Hecho en casa")
      expect(product.properties.first.presentation).to eql("Sí")
      expect(product.properties.second.name).to eql("For real no fake")
      expect(product.properties.second.presentation).to eql("No")
      #Options
      ot1 = product.option_types.first
      expect(ot1.name).to eql("Sabor")
      expect(ot1.option_values.first.name).to eql("Nueces")
      expect(ot1.option_values.second.name).to eql("Vainilla")
      ot2 = product.option_types.second
      expect(ot2.name).to eql("Tamaño")
      expect(ot2.option_values.first.name).to eql("estoy a dieta")
      expect(ot2.option_values.second.name).to eql("gigante")
      #variants
      var = product.variants.first
      expect(var.sku).to eql("12345678")
      expect(var.weight).to eql(200.0)
      expect(var.height).to eql(10.0)
      expect(var.width).to eql(10.0)
      expect(var.depth).to eql(10.0)
      expect(var.price).to eql(2000.0)
      val1 = var.option_values.first
      expect(val1.name).to eql("Nueces")
      val2 = var.option_values.second
      expect(val2.name).to eql("estoy a dieta")
      #Locations
      loc = Spree::StockLocation.where(admin_name: "Central").first
      expect(loc.name).to eql("Isla Diamante")
      expect(loc.address1).to eql("Playa 123")
      expect(loc.city).to eql("Til Til")
      expect(loc.address2).to eql("Juan algo 234")
      expect(loc.zipcode).to eql("12345")
      expect(loc.phone).to eql("76543469")
      expect(loc.country.name).to eql("Chile")
      expect(loc.country.iso_name).to eql("CHILE")
      expect(loc.state.name).to eql("Región Metropolitana")
      expect(loc.state.country.name).to eql("Chile")
      expect(loc.active).to eql(true)
      expect(loc.default).to eql(false)
      expect(loc.backorderable_default).to eql(true)
      expect(loc.propagate_all_variants).to eql(true)
      #Stock
      var = product.variants.first
      item = var.stock_items.first
      expect(item.stock_location_id).to eql(loc.id)
      expect(item.backorderable).to eql(true)
      expect(item.count_on_hand).to eql(20)
    end
  end
end
