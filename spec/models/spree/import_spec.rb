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
      var = Spree::Variant.where(product_id: product.id, is_master: true).first
      #Products
      expect(product.name).to eql("queque")
      expect(product.description).to eql("en molde de cupcake")
      expect(product.price).to eql(1500)
      expect(product.meta_description).to eql("Este es el mejor queque de Chile")
      expect(product.available_on.strftime("%Y-%m-%d")).to eql(DateTime.parse("2017-12-06 17:26:02 UTC").strftime("%Y-%m-%d"))
      expect(product.shipping_category_id).to eql(sc.id)
      expect(var.sku).to eql("87654321")
      expect(var.weight).to eql(200.0)
      expect(var.height).to eql(10.0)
      expect(var.width).to eql(10.0)
      expect(var.depth).to eql(10.0)
      expect(var.price).to eql(1500.0)
      #Propiedades
      expect(product.properties.first.name).to eql("Hecho en casa")
      expect(product.properties.first.presentation).to eql("Hecho en casa")
      expect(product.properties.second.name).to eql("For real no fake")
      expect(product.properties.second.presentation).to eql("For real no fake")
      pro_pro1 = product.properties.first.product_properties.first
      expect(pro_pro1.value).to eql("Sí")
      pro_pro2 = product.properties.second.product_properties.first
      expect(pro_pro2.value).to eql("No")
      #Options
      ot1 = product.option_types.first
      expect(ot1.name).to eql("Sabor")
      expect(ot1.option_values.first.name).to eql("Nueces")
      expect(ot1.option_values.second.name).to eql("Vainilla")
      ot2 = product.option_types.second
      expect(ot2.name).to eql("Tamaño")
      expect(ot2.option_values.first.name).to eql("Estoy a dieta")
      expect(ot2.option_values.second.name).to eql("Gigante")
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
      expect(val2.name).to eql("Estoy a dieta")
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
    it "should change master variant if product data is distinct" do
      #We load the excel
      TiendappProducts::Import.create_products('spec/fixtures/imported.xlsx')

      #Check the database for the correct entries
      product = Spree::Product.where(slug: "queque-en-molde-de-cupcake").first
      vr = Spree::Variant.where(product_id: product.id, is_master: true).first
      expect(vr.sku).to eql("87654321")
      expect(vr.weight).to eql(200.0)
      expect(vr.height).to eql(10.0)
      expect(vr.width).to eql(10.0)
      expect(vr.depth).to eql(10.0)
    end
    it "should create the correct taxonomies (and taxons)" do
      #We load the excel
      TiendappProducts::Import.create_products('spec/fixtures/imported.xlsx')

      #Check the database for the correct entries
      product = Spree::Product.where(slug: "queque-en-molde-de-cupcake").first
      taxons = product.taxons
      t1 = taxons.first
      ty = Spree::Taxonomy.find(t1.taxonomy_id)
      expect(t1.name).to eql("rere")
      expect(ty.name).to eql("riri")
      t2 = Spree::Taxon.find(t1.parent_id)
      expect(t2.name).to eql("ruru")
      expect(t2.taxonomy_id).to eql(ty.id)
      t3 = Spree::Taxon.find(t2.parent_id)
      expect(t3.name).to eql("riri")
      expect(t3.taxonomy_id).to eql(ty.id)
      t1 = taxons.second
      ty = Spree::Taxonomy.find(t1.taxonomy_id)
      expect(t1.name).to eql("coco")
      expect(ty.name).to eql("coco")
    end
    it "should still work if only the require data is in the excel" do
      #We load the excel
      TiendappProducts::Import.create_products('spec/fixtures/imported2.xlsx')

      # Product
      product = Spree::Product.where(slug: "queque").first
      sc = Spree::ShippingCategory.where(name: "Por defecto").first
      expect(product.name).to eql("queque")
      expect(product.description).to eql(nil)
      expect(product.price).to eql(1500)
      expect(product.meta_description).to eql(nil)
      expect(product.available_on).to eql(nil)
      expect(product.shipping_category_id).to eql(sc.id)
      expect(product.available?).to eql(false)
      #should only add SKU, weight, height, width and depth if there are not N/A or empty
      pr = Spree::Product.where(slug: "queque").first
      vr = Spree::Variant.where(product_id: pr.id, is_master: true).first
      expect(vr.sku).to eql("")
      expect(vr.weight).to eql(0.0)
      expect(vr.height).to eql(nil)
      expect(vr.width).to eql(nil)
      expect(vr.depth).to eql(nil)
      #should not have categories (taxons)
      expect(product.taxons.count).to eql(0)

      #Variants
      var = product.variants.first
      expect(var.sku).to eql("")
      expect(var.weight).to eql(0.0)
      expect(var.height).to eql(nil)
      expect(var.width).to eql(nil)
      expect(var.depth).to eql(nil)
      expect(var.price).to eql(2000.0)
      val1 = var.option_values.first
      expect(val1.name).to eql("Nueces")
      val2 = var.option_values.second
      expect(val2.name).to eql("Estoy a dieta")

      #Locations
      loc = Spree::StockLocation.where(admin_name: "Central").first
      expect(loc.name).to eql("Isla Diamante")
      expect(loc.address1).to eql("Playa 123")
      expect(loc.city).to eql("Til Til")
      expect(loc.address2).to eql("")
      expect(loc.zipcode).to eql("12345")
      expect(loc.phone).to eql("")
      expect(loc.country.name).to eql("Chile")
      expect(loc.country.iso_name).to eql("CHILE")
      expect(loc.state.name).to eql("Región Metropolitana")
      expect(loc.state.country.name).to eql("Chile")
      expect(loc.active).to eql(true)
      expect(loc.default).to eql(false)
      expect(loc.backorderable_default).to eql(false)
      expect(loc.propagate_all_variants).to eql(true)

      #Stock
      var = product.variants.first
      item = var.stock_items.first
      expect(item.stock_location_id).to eql(Spree::StockLocation.first.id)
      expect(item.backorderable).to eql(false)
      expect(item.count_on_hand).to eql(20)

    end
    it "should not add existing entries to database" do
      #We load the same excel twice
      TiendappProducts::Import.create_products('spec/fixtures/imported.xlsx')
      TiendappProducts::Import.create_products('spec/fixtures/imported.xlsx')

      expect(Spree::Product.count).to eql(2)
      pr = Spree::Product.first
      expect(pr.taxons.count).to eql(2)
      expect(Spree::Taxonomy.all.count).to eql(2)
      expect(Spree::Taxon.all.count).to eql(4)
      expect(pr.properties.count).to eql(2)
      expect(pr.properties.first.product_properties.count).to eql(1)
      expect(Spree::Property.count).to eql(2)
      expect(Spree::ProductProperty.count).to eql(2)
      expect(Spree::OptionType.count).to eql(2)
      expect(Spree::ProductOptionType.count).to eql(2)
      expect(Spree::OptionValue.count).to eql(4)
      expect(pr.option_types.count).to eql(2)
      expect(pr.option_types.first.option_values.count).to eql(2)
      expect(pr.option_types.second.option_values.count).to eql(2)
      expect(pr.variants.count).to eql(1)
      expect(Spree::Variant.all.count).to eql(3)
      expect(Spree::Country.count).to eql(1)
      expect(Spree::State.count).to eql(1)
      expect(Spree::StockLocation.count).to eql(1)
      expect(Spree::StockItem.count).to eql(3)
      expect(Spree::StockMovement.count).to eql(2)
    end
    it "should have the last stock cantity in count_on_hand" do
      #We load the same excel twice
      TiendappProducts::Import.create_products('spec/fixtures/imported.xlsx')
      TiendappProducts::Import.create_products('spec/fixtures/imported3.xlsx') #Diferent stock value

      product = Spree::Product.where(slug: "queque-en-molde-de-cupcake").first
      expect(product.stock_items.second.count_on_hand).to eql(10)
    end
    it "should return error if a require is missing in Products" do
      m = TiendappProducts::Import.create_products('spec/fixtures/imported4.xlsx')
      expect(m).to eql("Error de formato: en la hoja Productos fila 2 el Nombre no puede ser vacio")
    end
    it "should return error if a require is missing in Propiedades" do
      m = TiendappProducts::Import.create_products('spec/fixtures/imported5.xlsx')
      expect(m).to eql("Error de formato: en la hoja Propiedades fila 2 la Propiedad no puede ser vacio")
    end
    it "should return error if a require is missing in Opciones" do
      m = TiendappProducts::Import.create_products('spec/fixtures/imported6.xlsx')
      expect(m).to eql("Error de formato: en la hoja Opciones fila 2 los Valores no puede ser vacio")
    end
    it "should return error if a require is missing in Variantes" do
      m = TiendappProducts::Import.create_products('spec/fixtures/imported7.xlsx')
      expect(m).to eql("Error de formato: en la hoja Variantes fila 2 el Precio no puede ser vacio")
    end
    it "should return error if a require is missing in Ubicaciones" do
      m = TiendappProducts::Import.create_products('spec/fixtures/imported8.xlsx')
      expect(m).to eql("Error de formato: en la hoja Ubicaciones fila 2 la Calle no puede ser vacio")
    end
    it "should return error if a require is missing in Stock" do
      m = TiendappProducts::Import.create_products('spec/fixtures/imported9.xlsx')
      expect(m).to eql("Error de formato: en la hoja Variantes fila 2 el ID Variante no puede ser vacio")
    end
    it "should fail if the file to import is not a xlsx" do
      m = TiendappProducts::Import.create_products('spec/fixtures/imported.xls')
      expect(m).to eql("Error: la extensión del archivo debe ser xlsx")
    end
  end
end
