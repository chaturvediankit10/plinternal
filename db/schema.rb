# This file is auto-generated from the current state of the database. Instead
# of editing this file, please use the migrations feature of Active Record to
# incrementally modify your database, and then regenerate this schema definition.
#
# Note that this schema.rb definition is the authoritative source for your
# database schema. If you need to create the application database on another
# system, you should be using db:schema:load, not running all the migrations
# from scratch. The latter is a flawed and unsustainable approach (the more migrations
# you'll amass, the slower it'll run and the greater likelihood for issues).
#
# It's strongly recommended that you check this file into your version control system.

ActiveRecord::Schema.define(version: 2019_05_21_131350) do

  # These are extensions that must be enabled in order to support this database
  enable_extension "plpgsql"

  create_table "active_storage_attachments", force: :cascade do |t|
    t.string "name", null: false
    t.string "record_type", null: false
    t.bigint "record_id", null: false
    t.bigint "blob_id", null: false
    t.datetime "created_at", null: false
    t.index ["blob_id"], name: "index_active_storage_attachments_on_blob_id"
    t.index ["record_type", "record_id", "name", "blob_id"], name: "index_active_storage_attachments_uniqueness", unique: true
  end

  create_table "active_storage_blobs", force: :cascade do |t|
    t.string "key", null: false
    t.string "filename", null: false
    t.string "content_type"
    t.text "metadata"
    t.bigint "byte_size", null: false
    t.string "checksum", null: false
    t.datetime "created_at", null: false
    t.index ["key"], name: "index_active_storage_blobs_on_key", unique: true
  end

  create_table "adjustments", force: :cascade do |t|
    t.json "data"
    t.string "program_title"
    t.string "loan_category"
    t.integer "program_ids", default: [], array: true
    t.integer "program_id"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
    t.string "cell_number"
  end

  create_table "banks", force: :cascade do |t|
    t.string "name"
    t.integer "nmls"
    t.string "phone"
    t.string "address1"
    t.string "address2"
    t.string "city"
    t.string "state_code"
    t.string "zip"
    t.string "state_eligibility"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "error_logs", force: :cascade do |t|
    t.text "details"
    t.integer "column"
    t.integer "row"
    t.string "loan_category"
    t.integer "sheet_id"
    t.boolean "status", default: false
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
    t.string "bank_name"
    t.text "error_detail"
  end

  create_table "program_adjustments", force: :cascade do |t|
    t.integer "program_id"
    t.integer "adjustment_id"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "programs", force: :cascade do |t|
    t.integer "bank_id"
    t.integer "term"
    t.boolean "conforming", default: false
    t.boolean "fannie_mae", default: false
    t.boolean "fannie_mae_home_ready", default: false
    t.boolean "freddie_mac", default: false
    t.boolean "freddie_mac_home_possible", default: false
    t.boolean "fha", default: false
    t.boolean "va", default: false
    t.boolean "usda", default: false
    t.boolean "streamline", default: false
    t.boolean "full_doc", default: false
    t.text "adjustment_ids"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
    t.string "loan_category"
    t.json "base_rate"
    t.string "program_category"
    t.string "bank_name"
    t.string "program_name"
    t.string "rate_type"
    t.integer "sheet_id"
    t.string "loan_type"
    t.integer "lock_period", default: [], array: true
    t.string "loan_limit_type", default: [], array: true
    t.string "loan_purpose"
    t.string "arm_basic"
    t.string "arm_advanced"
    t.string "loan_size"
    t.string "fannie_mae_product"
    t.string "freddie_mac_product"
    t.integer "sub_sheet_id"
    t.boolean "fannie_mae_du"
    t.boolean "freddie_mac_lp"
    t.string "arm_benchmark"
    t.float "arm_margin"
    t.string "arm_caps"
  end

  create_table "sheets", force: :cascade do |t|
    t.string "name"
    t.integer "bank_id"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "sub_sheets", force: :cascade do |t|
    t.string "name"
    t.integer "sheet_id"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  create_table "user_inputs", force: :cascade do |t|
    t.text "property_type", default: [], array: true
    t.text "financing_type", default: [], array: true
    t.text "premium_type", default: [], array: true
    t.string "ltv", default: [], array: true
    t.string "fico", default: [], array: true
    t.text "refinance_option", default: [], array: true
    t.text "misc_adjuster", default: [], array: true
    t.boolean "lpmi"
    t.integer "coverage"
    t.integer "loan_amount"
    t.string "cltv"
    t.boolean "dti"
    t.float "interest_rate"
    t.integer "lock_period"
    t.string "state"
    t.datetime "created_at", null: false
    t.datetime "updated_at", null: false
  end

  add_foreign_key "active_storage_attachments", "active_storage_blobs", column: "blob_id"
end
