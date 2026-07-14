import React, { useState, useEffect, useMemo, useCallback } from "react";
import { Package, LogIn, LogOut, LayoutDashboard, ArrowDownToLine, ArrowUpFromLine, ClipboardList, Users, ScrollText, Search, Plus, Trash2, Pencil, X, KeyRound, ShieldCheck, Eye, AlertTriangle, Menu, CalendarRange, Download, Lock, ChevronDown, Truck } from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "./lib/supabaseClient.js";

/* ============================== MASTER DATA (seeded from SPPG Sleman Tempel Tambakrejo, Periode 16-28 Feb 2026) ============================== */
const ITEMS = [{"kode": "KH.01.001", "nama": "Beras putih (premium)", "satuan": "kg"}, {"kode": "KH.01.002", "nama": "Beras putih (medium)", "satuan": "kg"}, {"kode": "KH.01.003", "nama": "Beras merah", "satuan": "kg"}, {"kode": "KH.01.004", "nama": "Beras hitam", "satuan": "kg"}, {"kode": "KH.01.005", "nama": "Beras ketan putih", "satuan": "kg"}, {"kode": "KH.01.006", "nama": "Beras ketan hitam", "satuan": "kg"}, {"kode": "KH.01.007", "nama": "Tepung beras", "satuan": "kg"}, {"kode": "KH.01.008", "nama": "Tepung ketan", "satuan": "kg"}, {"kode": "KH.01.009", "nama": "Mie beras", "satuan": "kg"}, {"kode": "KH.01.099", "nama": "Beras dan olahan padi-lainnya", "satuan": "kg"}, {"kode": "KH.02.001", "nama": "Jagung pipil kering", "satuan": "kg"}, {"kode": "KH.02.002", "nama": "Jagung pipil segar", "satuan": "kg"}, {"kode": "KH.02.003", "nama": "Jagung manis segar", "satuan": "kg"}, {"kode": "KH.02.004", "nama": "Beras jagung", "satuan": "kg"}, {"kode": "KH.02.005", "nama": "Jagung giling", "satuan": "kg"}, {"kode": "KH.02.006", "nama": "Tepung jagung (maizena)", "satuan": "kg"}, {"kode": "KH.02.007", "nama": "Tepung jagung lokal (cornmeal)", "satuan": "kg"}, {"kode": "KH.02.099", "nama": "Jagung dan olahan-lainnya", "satuan": "kg"}, {"kode": "KH.03.001", "nama": "Singkong segar", "satuan": "kg"}, {"kode": "KH.03.002", "nama": "Singkong kering/gaplek", "satuan": "kg"}, {"kode": "KH.03.003", "nama": "Tepung singkong (tapioka)", "satuan": "kg"}, {"kode": "KH.03.004", "nama": "Mocaf", "satuan": "kg"}, {"kode": "KH.03.005", "nama": "Ubi jalar putih", "satuan": "kg"}, {"kode": "KH.03.006", "nama": "Ubi jalar kuning/oranye", "satuan": "kg"}, {"kode": "KH.03.007", "nama": "Ubi jalar ungu", "satuan": "kg"}, {"kode": "KH.03.008", "nama": "Kentang", "satuan": "kg"}, {"kode": "KH.03.009", "nama": "Talas", "satuan": "kg"}, {"kode": "KH.03.010", "nama": "Ganyong", "satuan": "kg"}, {"kode": "KH.03.011", "nama": "Sagu basah", "satuan": "kg"}, {"kode": "KH.03.012", "nama": "Tepung sagu", "satuan": "kg"}, {"kode": "KH.03.099", "nama": "Umbi umbian-lainnya", "satuan": "kg"}, {"kode": "KH.04.001", "nama": "Tepung terigu protein rendah", "satuan": "kg"}, {"kode": "KH.04.002", "nama": "Tepung terigu protein sedang", "satuan": "kg"}, {"kode": "KH.04.003", "nama": "Tepung terigu protein tinggi", "satuan": "kg"}, {"kode": "KH.04.004", "nama": "Tepung semolina", "satuan": "kg"}, {"kode": "KH.04.005", "nama": "Mie kering (non-instan)", "satuan": "pak"}, {"kode": "KH.04.006", "nama": "Spageti", "satuan": "pak"}, {"kode": "KH.04.007", "nama": "Makaroni", "satuan": "pak"}, {"kode": "KH.04.008", "nama": "Roti tawar", "satuan": "pak"}, {"kode": "KH.04.009", "nama": "Roti gandum", "satuan": "pak"}, {"kode": "KH.04.010", "nama": "Tepung panir", "satuan": "pak"}, {"kode": "KH.04.011", "nama": "Biskuit", "satuan": "pak"}, {"kode": "KH.04.012", "nama": "Bubur bayi", "satuan": "pak"}, {"kode": "KH.04.099", "nama": "Gandum dan olahan tepung-lainnya", "satuan": "pak"}, {"kode": "KH.05.001", "nama": "Oat (rolled oats)", "satuan": "kg"}, {"kode": "KH.05.002", "nama": "Oat instan", "satuan": "kg"}, {"kode": "KH.05.003", "nama": "Sorgum biji", "satuan": "kg"}, {"kode": "KH.05.004", "nama": "Tepung sorgum", "satuan": "kg"}, {"kode": "KH.05.005", "nama": "Millet/jewawut biji", "satuan": "kg"}, {"kode": "KH.05.006", "nama": "Tepung millet", "satuan": "kg"}, {"kode": "KH.05.099", "nama": "Sereal-lainnya", "satuan": "kg"}, {"kode": "PH.01.001", "nama": "Ayam broiler utuh", "satuan": "kg"}, {"kode": "PH.01.002", "nama": "Fillet dada ayam", "satuan": "kg"}, {"kode": "PH.01.003", "nama": "Paha ayam atas", "satuan": "kg"}, {"kode": "PH.01.004", "nama": "Paha ayam bawah (drumstick)", "satuan": "kg"}, {"kode": "PH.01.005", "nama": "Sayap ayam", "satuan": "kg"}, {"kode": "PH.01.006", "nama": "Hati ayam", "satuan": "kg"}, {"kode": "PH.01.007", "nama": "Ampela ayam", "satuan": "kg"}, {"kode": "PH.01.008", "nama": "Telur ayam ras/negeri", "satuan": "kg"}, {"kode": "PH.01.009", "nama": "Telur ayam kampung", "satuan": "kg"}, {"kode": "PH.01.010", "nama": "Telur puyuh", "satuan": "kg"}, {"kode": "PH.01.011", "nama": "Sosis ayam", "satuan": "pak"}, {"kode": "PH.01.099", "nama": "Ayam dan olahannya-lainnya", "satuan": "kg"}, {"kode": "PH.02.001", "nama": "Daging sapi segar (has luar/sirloin)", "satuan": "kg"}, {"kode": "PH.02.002", "nama": "Daging sapi segar (has dalam/tenderloin)", "satuan": "kg"}, {"kode": "PH.02.003", "nama": "Daging sapi segar (sandung lamur/brisket)", "satuan": "kg"}, {"kode": "PH.02.004", "nama": "Daging sapi segar (paha)", "satuan": "kg"}, {"kode": "PH.02.005", "nama": "Daging sapi giling", "satuan": "kg"}, {"kode": "PH.02.006", "nama": "Hati sapi", "satuan": "kg"}, {"kode": "PH.02.007", "nama": "Iga sapi", "satuan": "kg"}, {"kode": "PH.02.008", "nama": "Daging kerbau segar", "satuan": "kg"}, {"kode": "PH.02.009", "nama": "Daging kambing segar", "satuan": "kg"}, {"kode": "PH.02.010", "nama": "Daging domba segar", "satuan": "kg"}, {"kode": "PH.02.011", "nama": "Hati kambing", "satuan": "kg"}, {"kode": "PH.02.012", "nama": "Sosis sapi", "satuan": "pak"}, {"kode": "PH.02.099", "nama": "Daging sapi/kerbau/kambing-lainnya", "satuan": "kg"}, {"kode": "PH.03.001", "nama": "Lele", "satuan": "kg"}, {"kode": "PH.03.002", "nama": "Nila", "satuan": "kg"}, {"kode": "PH.03.003", "nama": "Mujair", "satuan": "kg"}, {"kode": "PH.03.004", "nama": "Patin", "satuan": "kg"}, {"kode": "PH.03.005", "nama": "Gurame", "satuan": "kg"}, {"kode": "PH.03.006", "nama": "Mas", "satuan": "kg"}, {"kode": "PH.03.007", "nama": "Gabus", "satuan": "kg"}, {"kode": "PH.03.008", "nama": "Fillet ikan tawar", "satuan": "kg"}, {"kode": "PH.03.099", "nama": "Ikan air tawar-lainnya", "satuan": "kg"}, {"kode": "PH.04.001", "nama": "Kembung", "satuan": "kg"}, {"kode": "PH.04.002", "nama": "Tongkol", "satuan": "kg"}, {"kode": "PH.04.003", "nama": "Cakalang", "satuan": "kg"}, {"kode": "PH.04.004", "nama": "Tuna", "satuan": "kg"}, {"kode": "PH.04.005", "nama": "Sarden/lemuru", "satuan": "kg"}, {"kode": "PH.04.006", "nama": "Bandeng", "satuan": "kg"}, {"kode": "PH.04.007", "nama": "Tenggiri", "satuan": "kg"}, {"kode": "PH.04.008", "nama": "Kakap", "satuan": "kg"}, {"kode": "PH.04.009", "nama": "Fillet ikan laut", "satuan": "kg"}, {"kode": "PH.04.010", "nama": "Ikan laut-lainnya", "satuan": "kg"}, {"kode": "PH.05.001", "nama": "Udang segar", "satuan": "kg"}, {"kode": "PH.05.002", "nama": "Udang beku", "satuan": "kg"}, {"kode": "PH.05.003", "nama": "Cumi-cumi", "satuan": "kg"}, {"kode": "PH.05.004", "nama": "Sotong", "satuan": "kg"}, {"kode": "PH.05.005", "nama": "Kepiting/rajungan", "satuan": "kg"}, {"kode": "PH.05.006", "nama": "Kerang hijau", "satuan": "kg"}, {"kode": "PH.05.007", "nama": "Kerang dara", "satuan": "kg"}, {"kode": "PH.05.099", "nama": "Produk perikanan lain-lainnya", "satuan": "kg"}, {"kode": "PH.06.001", "nama": "Susu UHT full cream", "satuan": "liter"}, {"kode": "PH.06.002", "nama": "Susu UHT rendah lemak", "satuan": "liter"}, {"kode": "PH.06.003", "nama": "Susu pasteurisasi", "satuan": "liter"}, {"kode": "PH.06.004", "nama": "Susu bubuk", "satuan": "kg"}, {"kode": "PH.06.005", "nama": "Yogurt plain", "satuan": "pak"}, {"kode": "PH.06.006", "nama": "Keju cheddar", "satuan": "pak"}, {"kode": "PH.06.007", "nama": "Keju olahan (slice)", "satuan": "pak"}, {"kode": "PH.06.099", "nama": "Susu dan olahan-lainnya", "satuan": "liter"}, {"kode": "PN.01.001", "nama": "Kedelai biji kering", "satuan": "kg"}, {"kode": "PN.01.002", "nama": "Tempe", "satuan": "kg"}, {"kode": "PN.01.003", "nama": "Tahu putih", "satuan": "kg"}, {"kode": "PN.01.004", "nama": "Tahu kuning", "satuan": "kg"}, {"kode": "PN.01.005", "nama": "Tahu sutra", "satuan": "kg"}, {"kode": "PN.01.006", "nama": "Susu kedelai", "satuan": "liter"}, {"kode": "PN.01.007", "nama": "Tepung kedelai", "satuan": "kg"}, {"kode": "PN.01.099", "nama": "Kedelai dan olahannya-lainnya", "satuan": "kg"}, {"kode": "PN.02.001", "nama": "Kacang tanah kupas", "satuan": "kg"}, {"kode": "PN.02.002", "nama": "Kacang tanah dengan kulit", "satuan": "kg"}, {"kode": "PN.02.003", "nama": "Kacang hijau", "satuan": "kg"}, {"kode": "PN.02.004", "nama": "Kacang merah", "satuan": "kg"}, {"kode": "PN.02.005", "nama": "Kacang tolo/tunggak", "satuan": "kg"}, {"kode": "PN.02.006", "nama": "Kacang hitam", "satuan": "kg"}, {"kode": "PN.02.007", "nama": "Kacang arab/chickpea", "satuan": "kg"}, {"kode": "PN.02.008", "nama": "Kacang polong/pea", "satuan": "kg"}, {"kode": "PN.02.009", "nama": "Kacang mede", "satuan": "kg"}, {"kode": "PN.02.010", "nama": "Kacang almond", "satuan": "kg"}, {"kode": "PN.02.099", "nama": "Kacang kacangan-lainnya", "satuan": "kg"}, {"kode": "PN.03.001", "nama": "Wijen", "satuan": "kg"}, {"kode": "PN.03.002", "nama": "Biji bunga matahari", "satuan": "kg"}, {"kode": "PN.03.003", "nama": "Chia seed", "satuan": "kg"}, {"kode": "PN.03.004", "nama": "Flaxseed", "satuan": "kg"}, {"kode": "PN.03.099", "nama": "Biji-bijian dan produknya-lainnya", "satuan": "kg"}, {"kode": "SY.01.001", "nama": "Bayam", "satuan": "ikat"}, {"kode": "SY.01.002", "nama": "Kangkung", "satuan": "ikat"}, {"kode": "SY.01.003", "nama": "Sawi hijau", "satuan": "ikat"}, {"kode": "SY.01.004", "nama": "Sawi putih", "satuan": "ikat"}, {"kode": "SY.01.005", "nama": "Pakcoy", "satuan": "ikat"}, {"kode": "SY.01.006", "nama": "Selada", "satuan": "ikat"}, {"kode": "SY.01.007", "nama": "Daun singkong", "satuan": "ikat"}, {"kode": "SY.01.008", "nama": "Daun pepaya", "satuan": "ikat"}, {"kode": "SY.01.009", "nama": "Kelor", "satuan": "ikat"}, {"kode": "SY.01.010", "nama": "Katuk", "satuan": "ikat"}, {"kode": "SY.01.011", "nama": "Seledri daun", "satuan": "ikat"}, {"kode": "SY.01.012", "nama": "Brokoli", "satuan": "kg"}, {"kode": "SY.01.013", "nama": "Kembang kol", "satuan": "kg"}, {"kode": "SY.01.014", "nama": "Buncis", "satuan": "kg"}, {"kode": "SY.01.015", "nama": "Seledri batang", "satuan": "ikat"}, {"kode": "SY.01.016", "nama": "Asparagus", "satuan": "ikat"}, {"kode": "SY.01.099", "nama": "Sayuran daun/batang/bunga-lainnya", "satuan": "ikat"}, {"kode": "SY.02.001", "nama": "Tomat", "satuan": "kg"}, {"kode": "SY.02.002", "nama": "Mentimun", "satuan": "kg"}, {"kode": "SY.02.003", "nama": "Terong ungu", "satuan": "kg"}, {"kode": "SY.02.004", "nama": "Terong hijau", "satuan": "kg"}, {"kode": "SY.02.005", "nama": "Cabai merah besar", "satuan": "kg"}, {"kode": "SY.02.006", "nama": "Cabai rawit", "satuan": "kg"}, {"kode": "SY.02.007", "nama": "Labu siam", "satuan": "kg"}, {"kode": "SY.02.008", "nama": "Pare", "satuan": "kg"}, {"kode": "SY.02.009", "nama": "Paprika", "satuan": "kg"}, {"kode": "SY.02.010", "nama": "Okra", "satuan": "kg"}, {"kode": "SY.02.011", "nama": "Wortel", "satuan": "kg"}, {"kode": "SY.02.012", "nama": "Lobak putih", "satuan": "kg"}, {"kode": "SY.02.013", "nama": "Bit", "satuan": "kg"}, {"kode": "SY.02.099", "nama": "Sayuran buah/umbi/akar-lainnya", "satuan": "kg"}, {"kode": "SY.03.001", "nama": "Bawang merah", "satuan": "kg"}, {"kode": "SY.03.002", "nama": "Bawang putih", "satuan": "kg"}, {"kode": "SY.03.003", "nama": "Bawang bombai", "satuan": "kg"}, {"kode": "SY.03.004", "nama": "Daun bawang", "satuan": "kg"}, {"kode": "SY.03.005", "nama": "Bawang prei", "satuan": "kg"}, {"kode": "SY.03.006", "nama": "Jahe", "satuan": "kg"}, {"kode": "SY.03.007", "nama": "Kunyit", "satuan": "kg"}, {"kode": "SY.03.008", "nama": "Lengkuas", "satuan": "kg"}, {"kode": "SY.03.009", "nama": "Kencur", "satuan": "kg"}, {"kode": "SY.03.010", "nama": "Serai", "satuan": "kg"}, {"kode": "SY.03.011", "nama": "Temu kunci", "satuan": "kg"}, {"kode": "SY.03.012", "nama": "Kemiri", "satuan": "kg"}, {"kode": "SY.03.099", "nama": "Sayuran bawang dan aromatik lainnya", "satuan": "kg"}, {"kode": "SY.04.001", "nama": "Jamur tiram", "satuan": "kg"}, {"kode": "SY.04.002", "nama": "Jamur kancing", "satuan": "kg"}, {"kode": "SY.04.003", "nama": "Jamur kuping", "satuan": "kg"}, {"kode": "SY.04.004", "nama": "Jamur shiitake", "satuan": "kg"}, {"kode": "SY.04.099", "nama": "Jamur-lainnya", "satuan": "kg"}, {"kode": "BU.01.001", "nama": "Pisang ambon", "satuan": "kg"}, {"kode": "BU.01.002", "nama": "Pisang kepok", "satuan": "kg"}, {"kode": "BU.01.003", "nama": "Pisang raja", "satuan": "kg"}, {"kode": "BU.01.004", "nama": "Pepaya", "satuan": "kg"}, {"kode": "BU.01.005", "nama": "Semangka", "satuan": "kg"}, {"kode": "BU.01.006", "nama": "Melon", "satuan": "kg"}, {"kode": "BU.01.007", "nama": "Mangga harum manis", "satuan": "kg"}, {"kode": "BU.01.008", "nama": "Mangga gedong", "satuan": "kg"}, {"kode": "BU.01.009", "nama": "Jeruk manis", "satuan": "kg"}, {"kode": "BU.01.010", "nama": "Jeruk keprok", "satuan": "kg"}, {"kode": "BU.01.011", "nama": "Jambu biji", "satuan": "kg"}, {"kode": "BU.01.012", "nama": "Jambu air", "satuan": "kg"}, {"kode": "BU.01.013", "nama": "Nanas", "satuan": "kg"}, {"kode": "BU.01.014", "nama": "Salak", "satuan": "kg"}, {"kode": "BU.01.015", "nama": "Rambutan", "satuan": "kg"}, {"kode": "BU.01.016", "nama": "Duku", "satuan": "kg"}, {"kode": "BU.01.017", "nama": "Sirsak", "satuan": "kg"}, {"kode": "BU.01.018", "nama": "Durian", "satuan": "kg"}, {"kode": "BU.01.019", "nama": "Manggis", "satuan": "kg"}, {"kode": "BU.01.020", "nama": "Alpukat", "satuan": "kg"}, {"kode": "BU.01.021", "nama": "Apel malang", "satuan": "kg"}, {"kode": "BU.01.022", "nama": "Buah naga merah", "satuan": "kg"}, {"kode": "BU.01.023", "nama": "Buah naga putih", "satuan": "kg"}, {"kode": "BU.01.024", "nama": "Kedondong", "satuan": "kg"}, {"kode": "BU.01.025", "nama": "Markisa", "satuan": "kg"}, {"kode": "BU.01.026", "nama": "Belimbing", "satuan": "kg"}, {"kode": "BU.01.027", "nama": "Sawo", "satuan": "kg"}, {"kode": "BU.01.028", "nama": "Matoa", "satuan": "kg"}, {"kode": "BU.01.099", "nama": "Buah lokal-lainnya", "satuan": "kg"}, {"kode": "BU.02.001", "nama": "Apel", "satuan": "kg"}, {"kode": "BU.02.002", "nama": "Pear", "satuan": "kg"}, {"kode": "BU.02.003", "nama": "Anggur merah", "satuan": "kg"}, {"kode": "BU.02.004", "nama": "Anggur hijau", "satuan": "kg"}, {"kode": "BU.02.005", "nama": "Kiwi", "satuan": "kg"}, {"kode": "BU.02.099", "nama": "Buah impor-lainnya", "satuan": "kg"}, {"kode": "BB.01.001", "nama": "Minyak goreng sawit", "satuan": "liter"}, {"kode": "BB.01.002", "nama": "Minyak goreng campuran", "satuan": "liter"}, {"kode": "BB.01.003", "nama": "Minyak kelapa", "satuan": "liter"}, {"kode": "BB.01.004", "nama": "Margarin", "satuan": "kg"}, {"kode": "BB.01.005", "nama": "Mentega (butter)", "satuan": "kg"}, {"kode": "BB.01.099", "nama": "Minyak dan lemak-lainnya", "satuan": "kg"}, {"kode": "BB.02.001", "nama": "Garam beryodium", "satuan": "pak"}, {"kode": "BB.02.002", "nama": "Gula pasir", "satuan": "kg"}, {"kode": "BB.02.003", "nama": "Gula merah/aren", "satuan": "kg"}, {"kode": "BB.02.004", "nama": "Madu", "satuan": "kg"}, {"kode": "BB.02.099", "nama": "Garam, gula, dan pemanis-lainnya", "satuan": "kg"}, {"kode": "BB.03.001", "nama": "Lada/merica bubuk", "satuan": "pcs"}, {"kode": "BB.03.002", "nama": "Ketumbar bubuk", "satuan": "pcs"}, {"kode": "BB.03.003", "nama": "Jinten bubuk", "satuan": "pcs"}, {"kode": "BB.03.004", "nama": "Pala bubuk", "satuan": "pcs"}, {"kode": "BB.03.005", "nama": "Kayu manis", "satuan": "pcs"}, {"kode": "BB.03.006", "nama": "Cengkeh", "satuan": "pcs"}, {"kode": "BB.03.007", "nama": "Kapulaga", "satuan": "pcs"}, {"kode": "BB.03.008", "nama": "Daun salam kering", "satuan": "pcs"}, {"kode": "BB.03.009", "nama": "Oregano", "satuan": "pcs"}, {"kode": "BB.03.099", "nama": "Bumbu kering (rempah)-lainnya", "satuan": "pcs"}, {"kode": "BB.04.001", "nama": "Kecap manis", "satuan": "botol"}, {"kode": "BB.04.002", "nama": "Kecap asin", "satuan": "botol"}, {"kode": "BB.04.003", "nama": "Saus sambal", "satuan": "botol"}, {"kode": "BB.04.004", "nama": "Saus tomat", "satuan": "botol"}, {"kode": "BB.04.005", "nama": "Saus tiram", "satuan": "botol"}, {"kode": "BB.04.006", "nama": "Cuka makan", "satuan": "botol"}, {"kode": "BB.04.007", "nama": "Mayones", "satuan": "botol"}, {"kode": "BB.04.099", "nama": "Bumbu cair, saus, dan kondimen-lainnya", "satuan": "botol"}, {"kode": "BB.05.001", "nama": "Santan segar", "satuan": "pcs"}, {"kode": "BB.05.002", "nama": "Santan instan cair", "satuan": "pcs"}, {"kode": "BB.05.003", "nama": "Santan bubuk", "satuan": "pcs"}, {"kode": "BB.05.004", "nama": "Kaldu bubuk ayam", "satuan": "pcs"}, {"kode": "BB.05.005", "nama": "Kaldu bubuk sapi", "satuan": "pcs"}, {"kode": "BB.05.006", "nama": "Kaldu jamur", "satuan": "pcs"}, {"kode": "BB.05.007", "nama": "Agar-agar bubuk", "satuan": "pcs"}, {"kode": "BB.05.008", "nama": "Gelatin", "satuan": "pcs"}, {"kode": "BB.05.099", "nama": "Produk olahan dasar-lainnya", "satuan": "pcs"}, {"kode": "BB.06.001", "nama": "Terasi", "satuan": "pcs"}, {"kode": "BB.06.002", "nama": "Tauco", "satuan": "pcs"}, {"kode": "BB.06.003", "nama": "Ragi roti", "satuan": "pcs"}, {"kode": "BB.06.004", "nama": "Baking powder", "satuan": "pcs"}, {"kode": "BB.06.099", "nama": "Bahan fermentasi/pelengkap-lainnya", "satuan": "pcs"}, {"kode": "BB.07.001", "nama": "Air minum kemasan galon", "satuan": "pcs"}, {"kode": "BB.07.002", "nama": "Air minum kemasan botol 2 liter", "satuan": "dus"}, {"kode": "BB.07.003", "nama": "Air minum kemasan botol 1,5 liter", "satuan": "dus"}, {"kode": "BB.07.004", "nama": "Air minum kemasan botol 600 ml", "satuan": "dus"}, {"kode": "BB.07.005", "nama": "Air minum kemasan botol 330 ml", "satuan": "dus"}, {"kode": "BB.07.006", "nama": "Air minum kemasan cup", "satuan": "dus"}, {"kode": "BB.07.007", "nama": "Minuman kemasan botol/kaleng", "satuan": "pcs"}, {"kode": "BB.07.008", "nama": "Susu kental manis", "satuan": "pcs"}, {"kode": "BB.07.099", "nama": "Bahan minuman pendamping-lainnya", "satuan": "pcs"}];
const SALDO_AWAL = {"KH.01.001": {"saldo": 12, "harga": 15000}, "KH.02.004": {"saldo": 8, "harga": 17000}, "KH.03.008": {"saldo": 5, "harga": 7000}, "KH.04.001": {"saldo": 7, "harga": 14000}, "KH.04.009": {"saldo": 15, "harga": 15000}, "KH.04.011": {"saldo": 5, "harga": 7500}, "KH.04.012": {"saldo": 4, "harga": 12500}, "PH.01.001": {"saldo": 5, "harga": 30000}, "PH.01.008": {"saldo": 10, "harga": 32000}, "PH.01.010": {"saldo": 5, "harga": 22000}, "PH.01.011": {"saldo": 6, "harga": 25000}, "PH.02.001": {"saldo": 10, "harga": 100000}, "PH.02.012": {"saldo": 7, "harga": 40000}, "PH.04.002": {"saldo": 4, "harga": 35000}, "PH.04.005": {"saldo": 5, "harga": 60000}, "PH.05.001": {"saldo": 5, "harga": 70000}, "PH.05.003": {"saldo": 9, "harga": 30000}, "PH.06.004": {"saldo": 5, "harga": 50000}, "PH.06.007": {"saldo": 18, "harga": 30000}, "PN.01.002": {"saldo": 3, "harga": 15000}, "PN.01.003": {"saldo": 2, "harga": 20000}, "PN.02.003": {"saldo": 17, "harga": 20000}, "PN.02.004": {"saldo": 9, "harga": 20000}, "BB.01.001": {"saldo": 9, "harga": 16000}};

const KELOMPOK = {
  KH: { nama: "Karbohidrat", color: "#C48A00", bg: "#FBF1DA" },
  PH: { nama: "Protein Hewani", color: "#A8402A", bg: "#F8E6E1" },
  PN: { nama: "Protein Nabati", color: "#6E7A3B", bg: "#EEF0DD" },
  SY: { nama: "Sayuran", color: "#2B6E52", bg: "#DEEEE6" },
  BU: { nama: "Buah-buahan", color: "#C1602E", bg: "#F8E4D6" },
  BB: { nama: "Bahan Baku Lain", color: "#5B6472", bg: "#E7E9EC" },
};

function kelompokOf(kode) { return kode.split(".")[0]; }

function fmtRp(n) {
  n = Number(n) || 0;
  return "Rp " + n.toLocaleString("id-ID");
}
function fmtDate(iso) {
  if (!iso) return "-";
  const d = new Date(iso);
  return d.toLocaleString("id-ID", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
}
function todayInput() {
  return new Date().toISOString().slice(0, 10);
}
function uid() { return Math.random().toString(36).slice(2, 10) + Date.now().toString(36); }

/* ============================== STORAGE HELPERS (Supabase relasional) ==============================
   Semua data (master barang, supplier, periode, saldo awal, transaksi masuk/keluar, log) disimpan
   sebagai baris-baris tabel Supabase sungguhan (lihat supabase/schema.sql), bukan blob JSON tunggal.
   Fungsi-fungsi mutasi di bawah melempar (throw) error apabila query gagal, supaya caller bisa
   menampilkan pesan yang jelas dan TIDAK diam-diam menimpa data asli dengan seed default (ini adalah
   bug data-loss yang pernah terjadi di versi app_kv sebelumnya).
   ================================================================================================= */

const DEFAULT_PERIODE_SEED = { nama: "Periode 1 - Februari 2026", mulai: "2026-02-16", selesai: "2026-02-28" };

function mapMasukRow(r) {
  return {
    id: r.id, tanggal: r.tanggal, supplier: r.supplier, kode: r.kode,
    nama: r.master_barang ? r.master_barang.nama : "", satuan: r.master_barang ? r.master_barang.satuan : "",
    vol: Number(r.vol), harga: Number(r.harga), jumlah: Number(r.jumlah),
    oleh: r.oleh, waktu: r.waktu, periodeId: r.periode_id,
  };
}
function mapKeluarRow(r) {
  return {
    id: r.id, tanggal: r.tanggal, kode: r.kode,
    nama: r.master_barang ? r.master_barang.nama : "", satuan: r.master_barang ? r.master_barang.satuan : "",
    vol: Number(r.vol), oleh: r.oleh, waktu: r.waktu, periodeId: r.periode_id,
  };
}
function mapLogRow(r) {
  return { id: r.id, waktu: r.waktu, oleh: r.oleh, role: r.role, aksi: r.aksi, detail: r.detail };
}
function mapPeriodeRow(p) {
  return { id: p.id, nama: p.nama, mulai: p.mulai, selesai: p.selesai, closed: p.closed };
}

function fmtTgl(iso) {
  if (!iso) return "-";
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("id-ID", { day: "2-digit", month: "short", year: "numeric" });
}

/* ============================== ROOT APP ============================== */
export default function App() {
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState(null);
  const [masuk, setMasuk] = useState([]);
  const [keluar, setKeluar] = useState([]);
  const [saldoByPeriode, setSaldoByPeriode] = useState({});
  const [periodeState, setPeriodeState] = useState({ list: [], activeId: null });
  const [masterItems, setMasterItems] = useState([]);
  const [suppliers, setSuppliers] = useState([]);
  const [log, setLog] = useState([]);
  const [session, setSession] = useState(null);
  const [page, setPage] = useState("dashboard");
  const [navOpen, setNavOpen] = useState(false);
  const [viewingPeriodeId, setViewingPeriodeId] = useState(null);

  // Dipakai untuk Admin/Super Admin: baca semua data dari tabel-tabel relasional (dibatasi RLS).
  // PENTING: setiap query di sini di-throw jika error, supaya kegagalan jaringan/RLS TIDAK pernah
  // diam-diam ditafsirkan sebagai "data kosong" (itulah bug data-loss yang pernah terjadi).
  const loadStaffData = useCallback(async () => {
    // 1) Master barang: seed dari ITEMS hanya jika tabel benar-benar kosong (bukan karena error).
    let { data: masterRows, error: masterErr } = await supabase.from("master_barang").select("*").order("kode");
    if (masterErr) throw masterErr;
    if (!masterRows || masterRows.length === 0) {
      const { error: seedErr } = await supabase.from("master_barang").insert(ITEMS);
      if (seedErr) throw seedErr;
      masterRows = ITEMS;
    }

    // 2) Periode: seed periode pertama + saldo awal dari SALDO_AWAL hanya jika tabel kosong.
    let { data: periodeRows, error: periodeErr } = await supabase.from("periode").select("*").order("mulai");
    if (periodeErr) throw periodeErr;
    if (!periodeRows || periodeRows.length === 0) {
      const seedP = DEFAULT_PERIODE_SEED;
      const { data: insertedP, error: insP } = await supabase.from("periode")
        .insert({ nama: seedP.nama, mulai: seedP.mulai, selesai: seedP.selesai, is_active: true }).select();
      if (insP) throw insP;
      periodeRows = insertedP;
      const saldoRows = Object.entries(SALDO_AWAL).map(([kode, v]) => ({ periode_id: insertedP[0].id, kode, saldo: v.saldo, harga: v.harga }));
      if (saldoRows.length) {
        const { error: seedSaldoErr } = await supabase.from("saldo_awal").insert(saldoRows);
        if (seedSaldoErr) throw seedSaldoErr;
      }
    }
    const activeRow = periodeRows.find((p) => p.is_active) || periodeRows[0];

    // 3) Sisanya: supplier, saldo_awal (semua periode), transaksi masuk/keluar, log.
    const [supplierRes, saldoRes, masukRes, keluarRes, logRes] = await Promise.all([
      supabase.from("supplier").select("*").order("nama"),
      supabase.from("saldo_awal").select("*"),
      supabase.from("transaksi_masuk").select("*, master_barang(nama, satuan)").order("waktu", { ascending: false }),
      supabase.from("transaksi_keluar").select("*, master_barang(nama, satuan)").order("waktu", { ascending: false }),
      supabase.from("activity_log").select("*").order("waktu", { ascending: false }).limit(500),
    ]);
    for (const r of [supplierRes, saldoRes, masukRes, keluarRes, logRes]) {
      if (r.error) throw r.error;
    }

    const saldoObj = {};
    for (const r of saldoRes.data || []) {
      if (!saldoObj[r.periode_id]) saldoObj[r.periode_id] = {};
      saldoObj[r.periode_id][r.kode] = { saldo: Number(r.saldo), harga: Number(r.harga) };
    }

    setMasterItems(masterRows);
    setSuppliers(supplierRes.data || []);
    setSaldoByPeriode(saldoObj);
    setMasuk((masukRes.data || []).map(mapMasukRow));
    setKeluar((keluarRes.data || []).map(mapKeluarRow));
    setLog((logRes.data || []).map(mapLogRow));
    setPeriodeState({ list: periodeRows.map(mapPeriodeRow), activeId: activeRow.id });
    setViewingPeriodeId(activeRow.id);
    setLoadError(null);
  }, []);

  // Pulihkan sesi Supabase Auth (Admin/Super Admin) kalau browser masih menyimpannya.
  useEffect(() => {
    (async () => {
      const { data: { session: authSession } } = await supabase.auth.getSession();
      if (authSession) {
        const { data: profile } = await supabase.from("profiles").select("*").eq("id", authSession.user.id).maybeSingle();
        if (profile) {
          setSession({ role: profile.role, nama: profile.nama, username: profile.username, id: profile.id });
          try {
            await loadStaffData();
          } catch (e) {
            setLoadError(e.message || String(e));
          }
        } else {
          await supabase.auth.signOut();
        }
      }
      setLoading(false);
    })();
  }, [loadStaffData]);

  async function handleLoginStaff(sess) {
    setLoading(true);
    setSession(sess);
    try {
      await loadStaffData();
    } catch (e) {
      setLoadError(e.message || String(e));
    }
    setLoading(false);
  }

  function handleLoginViewer(data) {
    setMasuk(data.masuk || []);
    setKeluar(data.keluar || []);
    setSaldoByPeriode(data.saldo || {});
    setMasterItems(data.master || []);
    setSuppliers(data.supplier || []);
    const pr = data.periode || { list: [], activeId: null };
    setPeriodeState(pr);
    setViewingPeriodeId(pr.activeId);
    setLog([]);
    setSession({ role: "viewer", nama: "Viewer", username: "viewer" });
  }

  async function handleLogout() {
    await supabase.auth.signOut();
    setSession(null);
    setPage("dashboard");
  }

  // Setiap fungsi di bawah menulis SATU baris ke tabel Supabase terkait (insert/update/delete),
  // lalu memperbarui state lokal dari hasil tulis itu saja — bukan menimpa seluruh tabel.
  const addLog = useCallback(async (aksi, detail) => {
    const { data, error } = await supabase.from("activity_log")
      .insert({ oleh: session ? session.nama : "-", role: session ? session.role : "-", aksi, detail })
      .select().single();
    if (error) { console.error("Gagal menyimpan log aktivitas", error); return; }
    setLog((prev) => [mapLogRow(data), ...prev].slice(0, 500));
  }, [session]);

  const persist = {
    tambahMasuk: async (entry) => {
      const { data, error } = await supabase.from("transaksi_masuk").insert({
        periode_id: entry.periodeId, tanggal: entry.tanggal, kode: entry.kode, supplier: entry.supplier,
        vol: entry.vol, harga: entry.harga, oleh: entry.oleh,
      }).select("*, master_barang(nama, satuan)").single();
      if (error) throw error;
      const mapped = mapMasukRow(data);
      setMasuk((prev) => [mapped, ...prev]);
      return mapped;
    },
    tambahKeluar: async (entry) => {
      const { data, error } = await supabase.from("transaksi_keluar").insert({
        periode_id: entry.periodeId, tanggal: entry.tanggal, kode: entry.kode, vol: entry.vol, oleh: entry.oleh,
      }).select("*, master_barang(nama, satuan)").single();
      if (error) throw error;
      const mapped = mapKeluarRow(data);
      setKeluar((prev) => [mapped, ...prev]);
      return mapped;
    },
    simpanMaster: async ({ kode, nama, satuan }, editKode) => {
      if (editKode) {
        const { error } = await supabase.from("master_barang").update({ nama, satuan }).eq("kode", editKode);
        if (error) throw error;
        setMasterItems((prev) => prev.map((i) => (i.kode === editKode ? { kode, nama, satuan } : i)));
      } else {
        const { error } = await supabase.from("master_barang").insert({ kode, nama, satuan });
        if (error) throw error;
        setMasterItems((prev) => [...prev, { kode, nama, satuan }]);
      }
    },
    hapusMaster: async (kode) => {
      const { error } = await supabase.from("master_barang").delete().eq("kode", kode);
      if (error) throw error;
      setMasterItems((prev) => prev.filter((i) => i.kode !== kode));
    },
    upsertSaldoAwal: async (periodeId, kode, { saldo, harga }) => {
      const { error } = await supabase.from("saldo_awal").upsert({ periode_id: periodeId, kode, saldo, harga });
      if (error) throw error;
      setSaldoByPeriode((prev) => ({ ...prev, [periodeId]: { ...(prev[periodeId] || {}), [kode]: { saldo, harga } } }));
    },
    simpanSupplier: async ({ id, nama, kontak }, editId) => {
      if (editId) {
        const { error } = await supabase.from("supplier").update({ nama, kontak }).eq("id", editId);
        if (error) throw error;
        setSuppliers((prev) => prev.map((s) => (s.id === editId ? { ...s, nama, kontak } : s)));
        return { id: editId, nama, kontak };
      }
      const { data, error } = await supabase.from("supplier").insert({ nama, kontak }).select().single();
      if (error) throw error;
      setSuppliers((prev) => [...prev, data]);
      return data;
    },
    hapusSupplier: async (id) => {
      const { error } = await supabase.from("supplier").delete().eq("id", id);
      if (error) throw error;
      setSuppliers((prev) => prev.filter((s) => s.id !== id));
    },
    editPeriodeInfo: async (id, patch) => {
      const { error } = await supabase.from("periode").update(patch).eq("id", id);
      if (error) throw error;
      setPeriodeState((prev) => ({ ...prev, list: prev.list.map((p) => (p.id === id ? { ...p, ...patch } : p)) }));
    },
  };

  // pastikan nama supplier tersimpan di data supplier; jika belum ada, tambahkan otomatis
  const ensureSupplier = useCallback(async (namaMentah) => {
    const nama = (namaMentah || "").trim();
    if (!nama) return "-";
    const existing = suppliers.find((s) => s.nama.toLowerCase() === nama.toLowerCase());
    if (existing) return existing.nama;
    await persist.simpanSupplier({ nama, kontak: "" }, null);
    await addLog("Tambah supplier baru (otomatis)", nama);
    return nama;
  }, [suppliers, addLog]);

  // hitung saldo akhir untuk sembarang periodeId (dipakai untuk tutup periode & laporan)
  const computeStockFor = useCallback((periodeId) => {
    const saldoAwalPeriode = saldoByPeriode[periodeId] || {};
    const map = {};
    for (const it of masterItems) {
      const sa = saldoAwalPeriode[it.kode] || {};
      map[it.kode] = { ...it, saldoAwal: sa.saldo || 0, hargaAwal: sa.harga || 0, masuk: 0, keluar: 0, hargaTerakhir: sa.harga || 0 };
    }
    for (const t of masuk) {
      if (t.periodeId !== periodeId || !map[t.kode]) continue;
      map[t.kode].masuk += Number(t.vol) || 0;
      if (t.harga) map[t.kode].hargaTerakhir = Number(t.harga);
    }
    for (const t of keluar) {
      if (t.periodeId !== periodeId || !map[t.kode]) continue;
      map[t.kode].keluar += Number(t.vol) || 0;
    }
    return Object.values(map).map((it) => {
      const saldoAkhir = it.saldoAwal + it.masuk - it.keluar;
      return { ...it, saldoAkhir, nilai: saldoAkhir * (it.hargaTerakhir || 0) };
    });
  }, [masterItems, saldoByPeriode, masuk, keluar]);

  const stock = useMemo(() => computeStockFor(viewingPeriodeId), [computeStockFor, viewingPeriodeId]);

  // Tutup periode aktif & buka periode baru lewat RPC atomik di database (lihat tutup_periode()
  // di schema.sql), lalu muat ulang semua data supaya state lokal konsisten dengan server.
  const tutupPeriode = useCallback(async ({ nama, mulai, selesai }) => {
    const namaLama = periodeState.list.find((p) => p.id === periodeState.activeId)?.nama;
    const { data: newId, error } = await supabase.rpc("tutup_periode", {
      p_nama_baru: nama, p_mulai_baru: mulai, p_selesai_baru: selesai,
    });
    if (error) throw error;
    await loadStaffData();
    setViewingPeriodeId(newId);
    await addLog("Tutup periode & buka periode baru", `${namaLama} -> ${nama}`);
  }, [periodeState, loadStaffData, addLog]);

  const editPeriode = useCallback(async (id, patch) => {
    await persist.editPeriodeInfo(id, patch);
    await addLog("Ubah data periode", `${id}`);
  }, [addLog]);

  if (loading) {
    return (
      <div style={{ minHeight: "100vh", background: "#F1EEE4", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", color: "#6B6355" }}>
        Memuat data gudang...
      </div>
    );
  }

  if (loadError) {
    return (
      <div style={{ minHeight: "100vh", background: "#F1EEE4", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", color: "#6B6355", padding: 20 }}>
        <div style={{ maxWidth: 420, textAlign: "center" }}>
          <AlertTriangle size={28} color="#A8402A" style={{ marginBottom: 10 }} />
          <div style={{ fontWeight: 700, marginBottom: 6, color: "#20241F" }}>Gagal memuat data dari Supabase</div>
          <div style={{ fontSize: 13, marginBottom: 16 }}>{loadError}</div>
          <div style={{ fontSize: 12, marginBottom: 16 }}>Untuk mencegah data hilang, aplikasi berhenti di sini alih-alih menampilkan data kosong. Periksa koneksi internet lalu coba lagi.</div>
          <button onClick={async () => { setLoading(true); try { await loadStaffData(); } catch (e) { setLoadError(e.message || String(e)); } setLoading(false); }} style={primaryBtn}>Coba Lagi</button>
        </div>
      </div>
    );
  }

  if (!session || !viewingPeriodeId) {
    return <LoginScreen onLoginStaff={handleLoginStaff} onLoginViewer={handleLoginViewer} />;
  }

  const periode = {
    list: periodeState.list,
    activeId: periodeState.activeId,
    viewingId: viewingPeriodeId,
    setViewingId: setViewingPeriodeId,
    tutupPeriode,
    editPeriode,
  };

  const ctx = { session, stock, masuk, keluar, masterItems, saldoByPeriode, log, persist, addLog, periode, computeStockFor, suppliers, ensureSupplier };

  return (
    <div style={{ minHeight: "100vh", background: "#F1EEE4", fontFamily: "Inter, system-ui, sans-serif", color: "#20241F" }}>
      <GlobalStyle />
      <TopBar session={session} onLogout={handleLogout} onMenu={() => setNavOpen((v) => !v)} />
      <PeriodeBar periode={periode} role={session.role} setPage={setPage} />
      {navOpen && <div className="nav-overlay show" onClick={() => setNavOpen(false)} />}
      <div style={{ display: "flex", maxWidth: 1280, margin: "0 auto" }}>
        <SideNav role={session.role} page={page} setPage={setPage} navOpen={navOpen} setNavOpen={setNavOpen} />
        <main style={{ flex: 1, padding: "20px 16px 60px", minWidth: 0 }}>
          <PageRouter page={page} ctx={ctx} />
        </main>
      </div>
    </div>
  );
}

function PeriodeBar({ periode, role, setPage }) {
  const current = periode.list.find((p) => p.id === periode.viewingId);
  const isActive = periode.viewingId === periode.activeId;
  const sorted = [...periode.list].sort((a, b) => (a.mulai < b.mulai ? 1 : -1));
  return (
    <div style={{ background: "#FBF9F3", borderBottom: "1px solid #E4DFCF", padding: "9px 16px", display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap", position: "sticky", top: 53, zIndex: 20 }}>
      <CalendarRange size={15} color="#2F5D50" style={{ flexShrink: 0 }} />
      <div style={{ position: "relative", flex: "1 1 220px", maxWidth: 320 }}>
        <select value={periode.viewingId} onChange={(e) => periode.setViewingId(e.target.value)}
          style={{ width: "100%", appearance: "none", padding: "7px 28px 7px 10px", borderRadius: 7, border: "1px solid #DAD4C2", background: "#fff", fontSize: 12.5, fontWeight: 600, color: "#2A2E27" }}>
          {sorted.map((p) => <option key={p.id} value={p.id}>{p.nama} ({fmtTgl(p.mulai)} - {fmtTgl(p.selesai)})</option>)}
        </select>
        <ChevronDown size={13} style={{ position: "absolute", right: 9, top: 10, color: "#8B8371", pointerEvents: "none" }} />
      </div>
      {isActive ? (
        <span style={{ fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.4, color: "#1B2E27", background: "#6FA88F", padding: "3px 8px", borderRadius: 5 }}>Periode Aktif</span>
      ) : (
        <span style={{ fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.4, color: "#fff", background: "#8B8371", padding: "3px 8px", borderRadius: 5, display: "flex", alignItems: "center", gap: 4 }}><Lock size={10} /> Arsip</span>
      )}
      {role === "superadmin" && (
        <button onClick={() => setPage("periode")} style={{ marginLeft: "auto", fontSize: 12, fontWeight: 600, color: "#2F5D50", background: "none", border: "none", cursor: "pointer", textDecoration: "underline", padding: "4px 2px" }}>
          Kelola Periode
        </button>
      )}
    </div>
  );
}

function GlobalStyle() {
  return (
    <style>{`
      @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');
      * { box-sizing: border-box; }
      body { margin: 0; }
      .sg { font-family: 'Space Grotesk', sans-serif; }
      .mono { font-family: 'IBM Plex Mono', monospace; }
      input, select, button, textarea { font-family: 'Inter', sans-serif; }
      input:focus, select:focus, textarea:focus, button:focus-visible { outline: 2px solid #2F5D50; outline-offset: 1px; }
      table { border-collapse: collapse; width: 100%; }
      ::-webkit-scrollbar { width: 8px; height: 8px; }
      ::-webkit-scrollbar-thumb { background: #D6CFBB; border-radius: 4px; }
      .menu-btn { display: none; }
      .nav-overlay { display: none; }
      .grid-2, .grid-master { display: grid; grid-template-columns: 380px 1fr; gap: 18px; align-items: flex-start; }
      .grid-master { grid-template-columns: 340px 1fr; }
      .grid-eq2 { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
      .grid-dash { display: grid; grid-template-columns: 1.15fr 1fr; gap: 16px; }
      .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
      .page-header-row { display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 18px; flex-wrap: wrap; gap: 10px; }
      @media (max-width: 900px) {
        .grid-2, .grid-master, .grid-eq2, .grid-dash { grid-template-columns: 1fr; }
      }
      @media (max-width: 820px) {
        .side-nav { position: fixed !important; z-index: 40; height: 100vh; top: 0; left: 0; transform: translateX(-100%); transition: transform .2s ease; box-shadow: 0 0 30px rgba(0,0,0,.2); }
        .side-nav.open { transform: translateX(0); }
        .menu-btn { display: flex; }
        .nav-overlay.show { display: block; position: fixed; inset: 0; background: rgba(20,25,20,.4); z-index: 35; }
      }
      @media (max-width: 640px) {
        .topbar-username { display: none; }
        table { font-size: 12px; }
        th, td { padding: 7px 6px !important; }
        .page-header-row { align-items: flex-start; flex-direction: column; }
        .page-header-row > *:last-child { width: 100%; }
      }
      @media (max-width: 480px) {
        h1.sg { font-size: 19px !important; }
      }
    `}</style>
  );
}

/* ============================== LOGIN ============================== */
function LoginScreen({ onLoginStaff, onLoginViewer }) {
  const [mode, setMode] = useState("staff");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  async function submitStaff(e) {
    e.preventDefault();
    setError("");
    if (!username || !password) { setError("Isi username dan password."); return; }
    setBusy(true);
    const email = `${username.trim().toLowerCase()}@sppg.local`;
    const { data, error: authErr } = await supabase.auth.signInWithPassword({ email, password });
    if (authErr || !data.user) {
      setBusy(false);
      setError("Username atau password salah.");
      return;
    }
    const { data: profile, error: profErr } = await supabase.from("profiles").select("*").eq("id", data.user.id).maybeSingle();
    if (profErr || !profile) {
      await supabase.auth.signOut();
      setBusy(false);
      setError("Akun ini belum terdaftar sebagai Admin/Super Admin. Hubungi Super Admin.");
      return;
    }
    setBusy(false);
    await onLoginStaff({ role: profile.role, nama: profile.nama, username: profile.username, id: profile.id });
  }

  async function submitPin(e) {
    e.preventDefault();
    setError("");
    if (!pin) { setError("Masukkan PIN."); return; }
    setBusy(true);
    const { data, error: rpcErr } = await supabase.rpc("get_viewer_data", { p_pin: pin });
    setBusy(false);
    if (rpcErr || !data) {
      setError("PIN salah.");
      return;
    }
    onLoginViewer(data);
  }

  return (
    <div style={{ minHeight: "100vh", background: "linear-gradient(160deg,#233A31,#152520)", display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: "Inter, sans-serif" }}>
      <GlobalStyle />
      <div style={{ width: "100%", maxWidth: 400 }}>
        <div style={{ textAlign: "center", marginBottom: 22, color: "#F1EEE4" }}>
          <div style={{ display: "inline-flex", alignItems: "center", justifyContent: "center", width: 52, height: 52, borderRadius: 14, background: "#E3A008", marginBottom: 14 }}>
            <Package size={26} color="#1B2E27" />
          </div>
          <div className="sg" style={{ fontSize: 22, fontWeight: 700, letterSpacing: -0.3 }}>Gudang Persediaan</div>
          <div style={{ fontSize: 13, color: "#B9C4BC", marginTop: 2 }}>SPPG · Bahan Pangan Dapur</div>
        </div>

        <div style={{ background: "#FBF9F3", borderRadius: 16, padding: 24, boxShadow: "0 20px 50px rgba(0,0,0,.35)" }}>
          <div style={{ display: "flex", gap: 4, background: "#EDE8DA", borderRadius: 10, padding: 4, marginBottom: 20 }}>
            <button onClick={() => { setMode("staff"); setError(""); }} style={{ flex: 1, padding: "9px 0", borderRadius: 7, border: "none", cursor: "pointer", fontWeight: 600, fontSize: 13.5, background: mode === "staff" ? "#2F5D50" : "transparent", color: mode === "staff" ? "#fff" : "#5C5646" }}>Admin / Super Admin</button>
            <button onClick={() => { setMode("viewer"); setError(""); }} style={{ flex: 1, padding: "9px 0", borderRadius: 7, border: "none", cursor: "pointer", fontWeight: 600, fontSize: 13.5, background: mode === "viewer" ? "#2F5D50" : "transparent", color: mode === "viewer" ? "#fff" : "#5C5646" }}>Viewer (PIN)</button>
          </div>

          {mode === "staff" ? (
            <form onSubmit={submitStaff}>
              <Field label="Username">
                <input value={username} onChange={(e) => setUsername(e.target.value)} autoFocus style={inputStyle} placeholder="mis. admin1" />
              </Field>
              <Field label="Password">
                <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} style={inputStyle} placeholder="********" />
              </Field>
              {error && <ErrorNote text={error} />}
              <button type="submit" disabled={busy} style={primaryBtn}><LogIn size={16} /> {busy ? "Memproses..." : "Masuk"}</button>
            </form>
          ) : (
            <form onSubmit={submitPin}>
              <Field label="Masukkan PIN">
                <input value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} inputMode="numeric" maxLength={8} autoFocus style={{ ...inputStyle, textAlign: "center", letterSpacing: 6, fontSize: 20 }} placeholder="----" />
              </Field>
              {error && <ErrorNote text={error} />}
              <button type="submit" disabled={busy} style={primaryBtn}><Eye size={16} /> {busy ? "Memproses..." : "Lihat Laporan"}</button>
            </form>
          )}
        </div>
      </div>
    </div>
  );
}

function ErrorNote({ text }) {
  return (
    <div style={{ display: "flex", gap: 6, alignItems: "center", background: "#F8E4E0", color: "#A8402A", padding: "8px 10px", borderRadius: 8, fontSize: 12.5, marginBottom: 12 }}>
      <AlertTriangle size={14} /> {text}
    </div>
  );
}
function Field({ label, children }) {
  return (
    <div style={{ marginBottom: 14 }}>
      <label style={{ display: "block", fontSize: 12, fontWeight: 600, color: "#5C5646", marginBottom: 5 }}>{label}</label>
      {children}
    </div>
  );
}
const inputStyle = { width: "100%", padding: "10px 12px", borderRadius: 8, border: "1px solid #DAD4C2", fontSize: 14, background: "#fff" };
const primaryBtn = { width: "100%", background: "#2F5D50", color: "#fff", border: "none", padding: "11px 0", borderRadius: 8, fontWeight: 600, fontSize: 14, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginTop: 4 };

/* ============================== TOP BAR / SIDE NAV ============================== */
function TopBar({ session, onLogout, onMenu }) {
  return (
    <div style={{ background: "#1B2E27", color: "#F1EEE4", padding: "10px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", position: "sticky", top: 0, zIndex: 30 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, minWidth: 0 }}>
        <button onClick={onMenu} className="menu-btn" style={{ background: "none", border: "none", color: "#F1EEE4", cursor: "pointer", padding: 4 }}><Menu size={20} /></button>
        <Package size={19} color="#E3A008" style={{ flexShrink: 0 }} />
        <span className="sg" style={{ fontWeight: 700, fontSize: 15, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>Gudang Persediaan</span>
        <span style={{ fontSize: 10.5, background: "#2F5D50", padding: "2px 7px", borderRadius: 20, marginLeft: 2, flexShrink: 0 }}>SPPG</span>
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 10, flexShrink: 0 }}>
        <RoleBadge role={session.role} />
        <span className="topbar-username" style={{ fontSize: 13 }}>{session.nama}</span>
        <button onClick={onLogout} title="Keluar" style={{ background: "rgba(255,255,255,.08)", border: "none", color: "#F1EEE4", padding: "6px 9px", borderRadius: 7, cursor: "pointer", display: "flex", alignItems: "center", gap: 6, fontSize: 12.5 }}>
          <LogOut size={14} />
        </button>
      </div>
    </div>
  );
}
function RoleBadge({ role }) {
  const map = { superadmin: { label: "Super Admin", color: "#E3A008" }, admin: { label: "Admin", color: "#6FA88F" }, viewer: { label: "Viewer", color: "#9AA4AE" } };
  const r = map[role];
  return <span style={{ fontSize: 10.5, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5, color: "#1B2E27", background: r.color, padding: "3px 8px", borderRadius: 5 }}>{r.label}</span>;
}

const NAV_ITEMS = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard, roles: ["superadmin", "admin", "viewer"] },
  { key: "masuk", label: "Penerimaan Barang", icon: ArrowDownToLine, roles: ["superadmin", "admin"] },
  { key: "keluar", label: "Pengeluaran Barang", icon: ArrowUpFromLine, roles: ["superadmin", "admin"] },
  { key: "laporan", label: "Laporan Stock", icon: ClipboardList, roles: ["superadmin", "admin", "viewer"] },
  { key: "master", label: "Master Barang", icon: Package, roles: ["superadmin"] },
  { key: "supplier", label: "Data Supplier", icon: Truck, roles: ["superadmin", "admin"] },
  { key: "periode", label: "Pengaturan Periode", icon: CalendarRange, roles: ["superadmin"] },
  { key: "users", label: "Manajemen User", icon: Users, roles: ["superadmin", "admin"] },
  { key: "log", label: "Log Aktivitas", icon: ScrollText, roles: ["superadmin"] },
];

function SideNav({ role, page, setPage, navOpen, setNavOpen }) {
  const items = NAV_ITEMS.filter((i) => i.roles.includes(role));
  return (
    <nav className={"side-nav" + (navOpen ? " open" : "")} style={{ width: 216, flexShrink: 0, padding: "20px 12px", background: "#FBF9F3", borderRight: "1px solid #E4DFCF" }}>
      {items.map((it) => {
        const Icon = it.icon;
        const active = page === it.key;
        return (
          <button key={it.key} onClick={() => { setPage(it.key); setNavOpen(false); }}
            style={{ width: "100%", display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", borderRadius: 9, border: "none", background: active ? "#2F5D50" : "transparent", color: active ? "#fff" : "#3A4038", fontWeight: active ? 600 : 500, fontSize: 13.5, cursor: "pointer", marginBottom: 3, textAlign: "left" }}>
            <Icon size={16} /> {it.label}
          </button>
        );
      })}
    </nav>
  );
}

function PageRouter({ page, ctx }) {
  const { session } = ctx;
  const found = NAV_ITEMS.find((n) => n.key === page);
  const allowed = found && found.roles.includes(session.role);
  if (!allowed) return <Dashboard ctx={ctx} />;
  switch (page) {
    case "dashboard": return <Dashboard ctx={ctx} />;
    case "masuk": return <Penerimaan ctx={ctx} />;
    case "keluar": return <Pengeluaran ctx={ctx} />;
    case "laporan": return <LaporanStock ctx={ctx} />;
    case "master": return <MasterBarang ctx={ctx} />;
    case "supplier": return <DataSupplier ctx={ctx} />;
    case "periode": return <PengaturanPeriode ctx={ctx} />;
    case "users": return <ManajemenUser ctx={ctx} />;
    case "log": return <LogAktivitas ctx={ctx} />;
    default: return <Dashboard ctx={ctx} />;
  }
}

/* ============================== SHARED UI ============================== */
function PageHeader({ title, subtitle, right }) {
  return (
    <div className="page-header-row">
      <div>
        <h1 className="sg" style={{ fontSize: 22, fontWeight: 700, margin: 0, color: "#1B2E27" }}>{title}</h1>
        {subtitle && <p style={{ margin: "4px 0 0", fontSize: 13, color: "#7A7362" }}>{subtitle}</p>}
      </div>
      {right}
    </div>
  );
}
function Card({ children, style }) {
  return <div style={{ background: "#FBF9F3", border: "1px solid #E4DFCF", borderRadius: 12, padding: 18, ...style }}>{children}</div>;
}
function KelompokChip({ kode }) {
  const k = KELOMPOK[kelompokOf(kode)];
  if (!k) return null;
  return <span className="mono" style={{ fontSize: 10.5, fontWeight: 600, color: k.color, background: k.bg, padding: "2px 7px", borderRadius: 5 }}>{kelompokOf(kode)}</span>;
}

/* ============================== DASHBOARD ============================== */
function Dashboard({ ctx }) {
  const { stock, masuk, keluar, session } = ctx;
  const totalNilai = stock.reduce((s, i) => s + i.nilai, 0);
  const lowStock = stock.filter((i) => i.saldoAkhir > 0 && i.saldoAkhir <= 3);
  const habis = stock.filter((i) => i.saldoAkhir <= 0);
  const recentMasuk = [...masuk].sort((a, b) => new Date(b.waktu) - new Date(a.waktu)).slice(0, 5);
  const recentKeluar = [...keluar].sort((a, b) => new Date(b.waktu) - new Date(a.waktu)).slice(0, 5);

  const byKelompok = useMemo(() => {
    const m = {};
    for (const it of stock) {
      const k = kelompokOf(it.kode);
      m[k] = (m[k] || 0) + it.nilai;
    }
    return m;
  }, [stock]);

  return (
    <div>
      <PageHeader title={`Halo, ${session.nama}`} subtitle="Ringkasan persediaan bahan pangan gudang." />
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(160px,1fr))", gap: 14, marginBottom: 20 }}>
        <StatCard label="Total Item Barang" value={stock.length} accent="#2F5D50" />
        <StatCard label="Nilai Persediaan" value={fmtRp(totalNilai)} accent="#E3A008" />
        <StatCard label="Stok Menipis (<=3)" value={lowStock.length} accent="#C1602E" />
        <StatCard label="Stok Habis" value={habis.length} accent="#A8402A" />
      </div>

      <div className="grid-dash">
        <Card>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 12, fontSize: 14.5 }}>Nilai Persediaan per Kelompok</div>
          {Object.entries(byKelompok).sort((a, b) => b[1] - a[1]).map(([k, v]) => {
            const info = KELOMPOK[k];
            const max = Math.max(...Object.values(byKelompok), 1);
            return (
              <div key={k} style={{ marginBottom: 10 }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 12.5, marginBottom: 4 }}>
                  <span style={{ fontWeight: 600 }}>{info ? info.nama : k}</span>
                  <span className="mono">{fmtRp(v)}</span>
                </div>
                <div style={{ height: 7, background: "#EDE8DA", borderRadius: 4, overflow: "hidden" }}>
                  <div style={{ width: `${(v / max) * 100}%`, height: "100%", background: info ? info.color : "#999" }} />
                </div>
              </div>
            );
          })}
        </Card>

        <Card>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 12, fontSize: 14.5, display: "flex", alignItems: "center", gap: 6 }}><AlertTriangle size={15} color="#C1602E" /> Perlu Perhatian</div>
          {habis.length === 0 && lowStock.length === 0 && <div style={{ fontSize: 13, color: "#8B8371" }}>Semua stok dalam kondisi aman.</div>}
          {[...habis, ...lowStock].slice(0, 8).map((it) => (
            <div key={it.kode} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: "1px solid #EEE9DB", fontSize: 12.5 }}>
              <span>{it.nama}</span>
              <span className="mono" style={{ color: it.saldoAkhir <= 0 ? "#A8402A" : "#C1602E", fontWeight: 600 }}>{it.saldoAkhir} {it.satuan}</span>
            </div>
          ))}
        </Card>
      </div>

      <div className="grid-eq2" style={{ marginTop: 16 }}>
        <Card>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10, fontSize: 14.5 }}>Penerimaan Terakhir</div>
          {recentMasuk.length === 0 && <Empty text="Belum ada transaksi penerimaan." />}
          {recentMasuk.map((t) => (
            <TxRow key={t.id} name={t.nama} sub={`${t.vol} ${t.satuan} - ${t.supplier || "-"}`} who={t.oleh} when={t.waktu} />
          ))}
        </Card>
        <Card>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10, fontSize: 14.5 }}>Pengeluaran Terakhir</div>
          {recentKeluar.length === 0 && <Empty text="Belum ada transaksi pengeluaran." />}
          {recentKeluar.map((t) => (
            <TxRow key={t.id} name={t.nama} sub={`${t.vol} ${t.satuan}`} who={t.oleh} when={t.waktu} />
          ))}
        </Card>
      </div>
    </div>
  );
}
function StatCard({ label, value, accent }) {
  return (
    <div style={{ background: "#FBF9F3", border: "1px solid #E4DFCF", borderTop: `3px solid ${accent}`, borderRadius: 10, padding: "14px 16px" }}>
      <div style={{ fontSize: 11.5, color: "#8B8371", fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4 }}>{label}</div>
      <div className="sg" style={{ fontSize: 20, fontWeight: 700, marginTop: 4 }}>{value}</div>
    </div>
  );
}
function TxRow({ name, sub, who, when }) {
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #EEE9DB" }}>
      <div>
        <div style={{ fontSize: 13, fontWeight: 500 }}>{name}</div>
        <div style={{ fontSize: 11.5, color: "#8B8371" }}>{sub}</div>
      </div>
      <div style={{ textAlign: "right" }}>
        <div style={{ fontSize: 11, color: "#5C5646" }}>{who}</div>
        <div style={{ fontSize: 10.5, color: "#A39B87" }}>{fmtDate(when)}</div>
      </div>
    </div>
  );
}
function Empty({ text }) {
  return <div style={{ fontSize: 12.5, color: "#8B8371", padding: "10px 0" }}>{text}</div>;
}

/* ============================== ITEM PICKER ============================== */
function ItemPicker({ items, value, onChange }) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const filtered = useMemo(() => {
    if (!q) return items.slice(0, 40);
    const qq = q.toLowerCase();
    return items.filter((i) => i.nama.toLowerCase().includes(qq) || i.kode.toLowerCase().includes(qq)).slice(0, 40);
  }, [q, items]);
  const selected = items.find((i) => i.kode === value);
  return (
    <div style={{ position: "relative" }}>
      <div style={{ position: "relative" }}>
        <Search size={14} style={{ position: "absolute", left: 10, top: 12, color: "#A39B87" }} />
        <input
          value={open ? q : selected ? selected.nama : ""}
          onChange={(e) => { setQ(e.target.value); setOpen(true); }}
          onFocus={() => { setOpen(true); setQ(""); }}
          placeholder="Cari nama atau kode barang..."
          style={{ ...inputStyle, paddingLeft: 30 }}
        />
      </div>
      {open && (
        <div style={{ position: "absolute", zIndex: 20, top: 42, left: 0, right: 0, maxHeight: 260, overflowY: "auto", background: "#fff", border: "1px solid #DAD4C2", borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,.12)" }}>
          {filtered.length === 0 && <div style={{ padding: 10, fontSize: 12.5, color: "#8B8371" }}>Tidak ditemukan.</div>}
          {filtered.map((it) => (
            <div key={it.kode} onClick={() => { onChange(it.kode); setOpen(false); setQ(""); }}
              style={{ padding: "8px 12px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid #F1EEE4", fontSize: 13 }}
              onMouseDown={(e) => e.preventDefault()}>
              <span>{it.nama} <span style={{ color: "#A39B87", fontSize: 11 }}>({it.satuan})</span></span>
              <KelompokChip kode={it.kode} />
            </div>
          ))}
        </div>
      )}
      {open && <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 10 }} />}
    </div>
  );
}

/* ============================== SUPPLIER PICKER ============================== */
function SupplierPicker({ suppliers, value, onChange }) {
  const [open, setOpen] = useState(false);
  const filtered = useMemo(() => {
    if (!value) return suppliers.slice(0, 30);
    const qq = value.toLowerCase();
    return suppliers.filter((s) => s.nama.toLowerCase().includes(qq)).slice(0, 30);
  }, [value, suppliers]);
  const exactMatch = suppliers.some((s) => s.nama.toLowerCase() === (value || "").trim().toLowerCase());

  return (
    <div style={{ position: "relative" }}>
      <div style={{ position: "relative" }}>
        <Truck size={14} style={{ position: "absolute", left: 10, top: 12, color: "#A39B87" }} />
        <input
          value={value}
          onChange={(e) => { onChange(e.target.value); setOpen(true); }}
          onFocus={() => setOpen(true)}
          placeholder="Ketik atau pilih nama supplier..."
          style={{ ...inputStyle, paddingLeft: 30 }}
        />
      </div>
      {open && (
        <div style={{ position: "absolute", zIndex: 20, top: 42, left: 0, right: 0, maxHeight: 220, overflowY: "auto", background: "#fff", border: "1px solid #DAD4C2", borderRadius: 8, boxShadow: "0 8px 24px rgba(0,0,0,.12)" }}>
          {filtered.map((s) => (
            <div key={s.id} onClick={() => { onChange(s.nama); setOpen(false); }} onMouseDown={(e) => e.preventDefault()}
              style={{ padding: "8px 12px", cursor: "pointer", borderBottom: "1px solid #F1EEE4", fontSize: 13 }}>
              {s.nama}
            </div>
          ))}
          {value && value.trim() && !exactMatch && (
            <div onClick={() => setOpen(false)} onMouseDown={(e) => e.preventDefault()}
              style={{ padding: "8px 12px", cursor: "pointer", fontSize: 12.5, color: "#2F5D50", fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
              <Plus size={13} /> Simpan sebagai supplier baru: "{value.trim()}"
            </div>
          )}
          {filtered.length === 0 && !(value && value.trim()) && <div style={{ padding: 10, fontSize: 12.5, color: "#8B8371" }}>Belum ada data supplier. Ketik nama untuk menambahkan.</div>}
        </div>
      )}
      {open && <div onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 10 }} />}
    </div>
  );
}

/* ============================== PENERIMAAN BARANG ============================== */
function Penerimaan({ ctx }) {
  const { masterItems, masuk, persist, session, addLog, periode, suppliers, ensureSupplier } = ctx;
  const isArsip = periode.viewingId !== periode.activeId;
  const [form, setForm] = useState({ tanggal: todayInput(), supplier: "", kode: "", vol: "", harga: "" });
  const [msg, setMsg] = useState("");

  async function submit(e) {
    e.preventDefault();
    const item = masterItems.find((i) => i.kode === form.kode);
    if (!item || !form.vol || Number(form.vol) <= 0) { setMsg("Lengkapi barang dan volume terlebih dahulu."); return; }
    const namaSupplier = await ensureSupplier(form.supplier);
    try {
      const entry = await persist.tambahMasuk({
        tanggal: form.tanggal, supplier: namaSupplier, kode: item.kode,
        vol: Number(form.vol), harga: Number(form.harga) || 0, oleh: session.nama, periodeId: periode.activeId,
      });
      await addLog("Penerimaan barang", `${item.nama} +${entry.vol} ${item.satuan}`);
      setForm({ tanggal: todayInput(), supplier: "", kode: "", vol: "", harga: "" });
      setMsg("Tersimpan.");
      setTimeout(() => setMsg(""), 2000);
    } catch (err) {
      setMsg("Gagal menyimpan ke Supabase: " + (err.message || String(err)));
    }
  }

  const mine = masuk.filter((t) => t.periodeId === periode.viewingId).sort((a, b) => new Date(b.waktu) - new Date(a.waktu));

  return (
    <div>
      <PageHeader title="Penerimaan Barang" subtitle="Catat barang yang diterima dari supplier. Tercatat otomatis atas nama akun yang login." />
      {isArsip && <ArsipNotice />}
      <div className="grid-2">
        {!isArsip && (
          <Card>
            <form onSubmit={submit}>
              <Field label="Tanggal"><input type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} style={inputStyle} /></Field>
              <Field label="Nama Supplier"><SupplierPicker suppliers={suppliers} value={form.supplier} onChange={(v) => setForm({ ...form, supplier: v })} /></Field>
              <Field label="Barang"><ItemPicker items={masterItems} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
              <div style={{ display: "flex", gap: 10 }}>
                <div style={{ flex: 1 }}><Field label="Volume"><input type="number" min="0" value={form.vol} onChange={(e) => setForm({ ...form, vol: e.target.value })} style={inputStyle} /></Field></div>
                <div style={{ flex: 1 }}><Field label="Harga Beli / satuan"><input type="number" min="0" value={form.harga} onChange={(e) => setForm({ ...form, harga: e.target.value })} style={inputStyle} /></Field></div>
              </div>
              {msg && <div style={{ fontSize: 12.5, color: "#2F5D50", marginBottom: 10 }}>{msg}</div>}
              <button type="submit" style={primaryBtn}><Plus size={16} /> Simpan Penerimaan</button>
            </form>
          </Card>
        )}
        <Card style={{ maxHeight: 520, overflowY: "auto" }}>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10 }}>Riwayat Penerimaan</div>
          <div className="table-wrap">
            <table>
              <thead><tr>{["Tanggal", "Barang", "Supplier", "Vol", "Harga", "Jumlah", "Oleh"].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
              <tbody>
                {mine.map((t) => (
                  <tr key={t.id}>
                    <Td>{t.tanggal}</Td>
                    <Td>{t.nama}</Td>
                    <Td>{t.supplier}</Td>
                    <Td className="mono">{t.vol} {t.satuan}</Td>
                    <Td className="mono">{fmtRp(t.harga)}</Td>
                    <Td className="mono">{fmtRp(t.jumlah)}</Td>
                    <Td>{t.oleh}<div style={{ fontSize: 10, color: "#A39B87" }}>{fmtDate(t.waktu)}</div></Td>
                  </tr>
                ))}
                {mine.length === 0 && <tr><td colSpan={7}><Empty text="Belum ada data." /></td></tr>}
              </tbody>
            </table>
          </div>
        </Card>
      </div>
    </div>
  );
}
function ArsipNotice() {
  return (
    <div style={{ display: "flex", gap: 8, alignItems: "center", background: "#EFEAD9", color: "#5C5646", padding: "10px 12px", borderRadius: 8, fontSize: 12.5, marginBottom: 16 }}>
      <Lock size={14} /> Anda sedang melihat periode arsip (tidak aktif). Input transaksi baru hanya berlaku pada periode yang sedang aktif — pilih periode aktif di bilah atas untuk mencatat transaksi.
    </div>
  );
}

/* ============================== PENGELUARAN BARANG ============================== */
function Pengeluaran({ ctx }) {
  const { masterItems, keluar, persist, session, addLog, stock, periode } = ctx;
  const isArsip = periode.viewingId !== periode.activeId;
  const [form, setForm] = useState({ tanggal: todayInput(), kode: "", vol: "" });
  const [msg, setMsg] = useState("");

  async function submit(e) {
    e.preventDefault();
    const item = masterItems.find((i) => i.kode === form.kode);
    if (!item || !form.vol || Number(form.vol) <= 0) { setMsg("Lengkapi barang dan volume terlebih dahulu."); return; }
    const st = stock.find((s) => s.kode === item.kode);
    if (st && Number(form.vol) > st.saldoAkhir) { setMsg(`Stok tidak cukup. Sisa stok: ${st.saldoAkhir} ${item.satuan}.`); return; }
    try {
      const entry = await persist.tambahKeluar({
        tanggal: form.tanggal, kode: item.kode, vol: Number(form.vol), oleh: session.nama, periodeId: periode.activeId,
      });
      await addLog("Pengeluaran barang", `${item.nama} -${entry.vol} ${item.satuan}`);
      setForm({ tanggal: todayInput(), kode: "", vol: "" });
      setMsg("Tersimpan.");
      setTimeout(() => setMsg(""), 2000);
    } catch (err) {
      setMsg("Gagal menyimpan ke Supabase: " + (err.message || String(err)));
    }
  }

  const mine = keluar.filter((t) => t.periodeId === periode.viewingId).sort((a, b) => new Date(b.waktu) - new Date(a.waktu));

  return (
    <div>
      <PageHeader title="Pengeluaran Barang" subtitle="Catat pemakaian bahan dari gudang. Tercatat otomatis atas nama akun yang login." />
      {isArsip && <ArsipNotice />}
      <div className="grid-2">
        {!isArsip && (
          <Card>
            <form onSubmit={submit}>
              <Field label="Tanggal"><input type="date" value={form.tanggal} onChange={(e) => setForm({ ...form, tanggal: e.target.value })} style={inputStyle} /></Field>
              <Field label="Barang"><ItemPicker items={masterItems} value={form.kode} onChange={(v) => setForm({ ...form, kode: v })} /></Field>
              <Field label="Volume"><input type="number" min="0" value={form.vol} onChange={(e) => setForm({ ...form, vol: e.target.value })} style={inputStyle} /></Field>
              {msg && <div style={{ fontSize: 12.5, color: msg.includes("cukup") ? "#A8402A" : "#2F5D50", marginBottom: 10 }}>{msg}</div>}
              <button type="submit" style={primaryBtn}><Plus size={16} /> Simpan Pengeluaran</button>
            </form>
          </Card>
        )}
        <Card style={{ maxHeight: 520, overflowY: "auto" }}>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10 }}>Riwayat Pengeluaran</div>
          <div className="table-wrap">
            <table>
              <thead><tr>{["Tanggal", "Barang", "Vol", "Petugas"].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
              <tbody>
                {mine.map((t) => (
                  <tr key={t.id}>
                    <Td>{t.tanggal}</Td>
                    <Td>{t.nama}</Td>
                    <Td className="mono">{t.vol} {t.satuan}</Td>
                    <Td>{t.oleh}<div style={{ fontSize: 10, color: "#A39B87" }}>{fmtDate(t.waktu)}</div></Td>
                  </tr>
                ))}
                {mine.length === 0 && <tr><td colSpan={4}><Empty text="Belum ada data." /></td></tr>}
              </tbody>
            </table>
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ============================== LAPORAN STOCK ============================== */
function LaporanStock({ ctx }) {
  const { stock, periode } = ctx;
  const [q, setQ] = useState("");
  const [kel, setKel] = useState("ALL");
  const periodeInfo = periode.list.find((p) => p.id === periode.viewingId);

  const rows = stock.filter((it) => {
    if (kel !== "ALL" && kelompokOf(it.kode) !== kel) return false;
    if (q && !(it.nama.toLowerCase().includes(q.toLowerCase()) || it.kode.toLowerCase().includes(q.toLowerCase()))) return false;
    return true;
  }).sort((a, b) => a.kode.localeCompare(b.kode));

  const rekap = useMemo(() => {
    const m = {};
    for (const it of stock) {
      const k = kelompokOf(it.kode);
      if (!m[k]) m[k] = { saldoAwal: 0, masuk: 0, keluar: 0, saldoAkhir: 0, nilai: 0 };
      m[k].saldoAwal += it.saldoAwal; m[k].masuk += it.masuk; m[k].keluar += it.keluar; m[k].saldoAkhir += it.saldoAkhir; m[k].nilai += it.nilai;
    }
    return m;
  }, [stock]);

  function exportExcel() {
    const wsRekap = XLSX.utils.json_to_sheet(
      Object.entries(rekap).map(([k, v]) => ({
        Kelompok: KELOMPOK[k] ? KELOMPOK[k].nama : k,
        "Saldo Awal": v.saldoAwal, Masuk: v.masuk, Keluar: v.keluar, "Saldo Akhir": v.saldoAkhir, "Nilai (Rp)": v.nilai,
      }))
    );
    const wsDetail = XLSX.utils.json_to_sheet(
      stock.slice().sort((a, b) => a.kode.localeCompare(b.kode)).map((it) => ({
        Kode: it.kode, "Nama Barang": it.nama, Satuan: it.satuan, "Saldo Awal": it.saldoAwal, Masuk: it.masuk, Keluar: it.keluar,
        "Saldo Akhir": it.saldoAkhir, "Harga Terakhir (Rp)": it.hargaTerakhir || 0, "Nilai (Rp)": it.nilai,
      }))
    );
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsRekap, "Rekap per Kelompok");
    XLSX.utils.book_append_sheet(wb, wsDetail, "Detail Barang");
    const namaFile = `Laporan_Stock_${(periodeInfo ? periodeInfo.nama : "Periode").replace(/[^a-zA-Z0-9]+/g, "_")}.xlsx`;
    XLSX.writeFile(wb, namaFile);
  }

  return (
    <div>
      <PageHeader
        title="Laporan Stock"
        subtitle={`Rekap dan rincian saldo barang persediaan - ${periodeInfo ? periodeInfo.nama : ""}.`}
        right={
          <button onClick={exportExcel} style={{ ...primaryBtn, width: "auto", padding: "9px 16px" }}>
            <Download size={15} /> Export Excel
          </button>
        }
      />

      <Card style={{ marginBottom: 16 }}>
        <div className="sg" style={{ fontWeight: 700, marginBottom: 10 }}>Rekap per Kelompok</div>
        <div className="table-wrap">
          <table>
            <thead><tr>{["Kelompok", "Saldo Awal", "Masuk", "Keluar", "Saldo Akhir", "Nilai"].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
            <tbody>
              {Object.entries(rekap).map(([k, v]) => (
                <tr key={k}>
                  <Td><KelompokChip kode={k + ".x"} /> {KELOMPOK[k] ? KELOMPOK[k].nama : k}</Td>
                  <Td className="mono">{v.saldoAwal}</Td>
                  <Td className="mono" style={{ color: "#2F5D50" }}>+{v.masuk}</Td>
                  <Td className="mono" style={{ color: "#A8402A" }}>-{v.keluar}</Td>
                  <Td className="mono" style={{ fontWeight: 700 }}>{v.saldoAkhir}</Td>
                  <Td className="mono">{fmtRp(v.nilai)}</Td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>

      <Card>
        <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
          <div style={{ position: "relative", flex: 1, minWidth: 200 }}>
            <Search size={14} style={{ position: "absolute", left: 10, top: 11, color: "#A39B87" }} />
            <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Cari barang..." style={{ ...inputStyle, paddingLeft: 30 }} />
          </div>
          <select value={kel} onChange={(e) => setKel(e.target.value)} style={{ ...inputStyle, width: 190 }}>
            <option value="ALL">Semua Kelompok</option>
            {Object.entries(KELOMPOK).map(([k, v]) => <option key={k} value={k}>{v.nama}</option>)}
          </select>
        </div>
        <div className="table-wrap">
          <table>
            <thead><tr>{["Kode", "Nama Barang", "Satuan", "Awal", "Masuk", "Keluar", "Akhir", "Nilai"].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
            <tbody>
              {rows.map((it) => (
                <tr key={it.kode}>
                  <Td className="mono">{it.kode}</Td>
                  <Td>{it.nama}</Td>
                  <Td>{it.satuan}</Td>
                  <Td className="mono">{it.saldoAwal}</Td>
                  <Td className="mono" style={{ color: "#2F5D50" }}>+{it.masuk}</Td>
                  <Td className="mono" style={{ color: "#A8402A" }}>-{it.keluar}</Td>
                  <Td className="mono" style={{ fontWeight: 700, color: it.saldoAkhir <= 0 ? "#A8402A" : it.saldoAkhir <= 3 ? "#C1602E" : "#20241F" }}>{it.saldoAkhir}</Td>
                  <Td className="mono">{fmtRp(it.nilai)}</Td>
                </tr>
              ))}
              {rows.length === 0 && <tr><td colSpan={8}><Empty text="Tidak ada barang yang cocok." /></td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ============================== MASTER BARANG (Super Admin) ============================== */
function MasterBarang({ ctx }) {
  const { masterItems, persist, addLog, saldoByPeriode, periode } = ctx;
  const saldoAwal = saldoByPeriode[periode.activeId] || {};
  const periodeAktifInfo = periode.list.find((p) => p.id === periode.activeId);
  const [q, setQ] = useState("");
  const [form, setForm] = useState({ kode: "", nama: "", satuan: "", saldo: "", harga: "" });
  const [editKode, setEditKode] = useState(null);
  const [msg, setMsg] = useState("");

  const filtered = masterItems.filter((i) => !q || i.nama.toLowerCase().includes(q.toLowerCase()) || i.kode.toLowerCase().includes(q.toLowerCase())).sort((a, b) => a.kode.localeCompare(b.kode));

  async function addItem(e) {
    e.preventDefault();
    if (!form.kode || !form.nama || !form.satuan) { setMsg("Kode, nama, dan satuan wajib diisi."); return; }
    if (masterItems.some((i) => i.kode === form.kode) && !editKode) { setMsg("Kode sudah dipakai."); return; }
    try {
      await persist.simpanMaster({ kode: form.kode, nama: form.nama, satuan: form.satuan }, editKode);
      if (form.saldo || form.harga) {
        await persist.upsertSaldoAwal(periode.activeId, form.kode, { saldo: Number(form.saldo) || 0, harga: Number(form.harga) || 0 });
      }
      await addLog(editKode ? "Ubah master barang" : "Tambah master barang", `${form.kode} - ${form.nama}`);
      setForm({ kode: "", nama: "", satuan: "", saldo: "", harga: "" });
      setEditKode(null);
      setMsg("Tersimpan.");
      setTimeout(() => setMsg(""), 2000);
    } catch (err) {
      setMsg("Gagal menyimpan ke Supabase: " + (err.message || String(err)));
    }
  }

  function editRow(it) {
    const sa = saldoAwal[it.kode] || {};
    setEditKode(it.kode);
    setForm({ kode: it.kode, nama: it.nama, satuan: it.satuan, saldo: sa.saldo || "", harga: sa.harga || "" });
  }

  async function removeItem(kode) {
    if (!confirm("Hapus barang ini dari master data?")) return;
    try {
      await persist.hapusMaster(kode);
      await addLog("Hapus master barang", kode);
    } catch (err) {
      alert("Barang tidak bisa dihapus, kemungkinan karena sudah punya riwayat transaksi masuk/keluar.\n\n" + (err.message || String(err)));
    }
  }

  return (
    <div>
      <PageHeader title="Master Barang" subtitle={`Kelola daftar barang persediaan. Saldo awal berlaku untuk periode aktif: ${periodeAktifInfo ? periodeAktifInfo.nama : "-"}.`} />
      <div className="grid-master">
        <Card>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10 }}>{editKode ? "Ubah Barang" : "Tambah Barang"}</div>
          <form onSubmit={addItem}>
            <Field label="Kode Barang"><input value={form.kode} onChange={(e) => setForm({ ...form, kode: e.target.value })} style={inputStyle} placeholder="mis. KH.01.010" disabled={!!editKode} /></Field>
            <Field label="Nama Barang"><input value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} style={inputStyle} /></Field>
            <Field label="Satuan"><input value={form.satuan} onChange={(e) => setForm({ ...form, satuan: e.target.value })} style={inputStyle} placeholder="kg / pcs / liter" /></Field>
            <div style={{ display: "flex", gap: 10 }}>
              <div style={{ flex: 1 }}><Field label={`Saldo Awal (${periodeAktifInfo ? periodeAktifInfo.nama : "periode aktif"})`}><input type="number" value={form.saldo} onChange={(e) => setForm({ ...form, saldo: e.target.value })} style={inputStyle} /></Field></div>
              <div style={{ flex: 1 }}><Field label="Harga Awal"><input type="number" value={form.harga} onChange={(e) => setForm({ ...form, harga: e.target.value })} style={inputStyle} /></Field></div>
            </div>
            {msg && <div style={{ fontSize: 12.5, marginBottom: 10, color: "#2F5D50" }}>{msg}</div>}
            <div style={{ display: "flex", gap: 8 }}>
              <button type="submit" style={primaryBtn}>{editKode ? <><Pencil size={15} /> Simpan Perubahan</> : <><Plus size={15} /> Tambah Barang</>}</button>
              {editKode && <button type="button" onClick={() => { setEditKode(null); setForm({ kode: "", nama: "", satuan: "", saldo: "", harga: "" }); }} style={{ ...primaryBtn, background: "#EDE8DA", color: "#3A4038" }}><X size={15} /></button>}
            </div>
          </form>
        </Card>
        <Card style={{ maxHeight: 560, overflowY: "auto" }}>
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Cari barang..." style={{ ...inputStyle, marginBottom: 10 }} />
          <div className="table-wrap">
            <table>
              <thead><tr>{["Kode", "Nama", "Satuan", "Saldo Awal", ""].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
              <tbody>
                {filtered.map((it) => (
                  <tr key={it.kode}>
                    <Td className="mono">{it.kode}</Td>
                    <Td>{it.nama}</Td>
                    <Td>{it.satuan}</Td>
                    <Td className="mono">{(saldoAwal[it.kode] && saldoAwal[it.kode].saldo) || 0}</Td>
                    <Td>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => editRow(it)} style={iconBtn}><Pencil size={13} /></button>
                        <button onClick={() => removeItem(it.kode)} style={{ ...iconBtn, color: "#A8402A" }}><Trash2 size={13} /></button>
                      </div>
                    </Td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Card>
      </div>
    </div>
  );
}
const iconBtn = { border: "1px solid #DAD4C2", background: "#fff", borderRadius: 6, padding: 5, cursor: "pointer", display: "flex" };

/* ============================== DATA SUPPLIER ============================== */
function DataSupplier({ ctx }) {
  const { suppliers, persist, addLog, masuk } = ctx;
  const [q, setQ] = useState("");
  const [form, setForm] = useState({ nama: "", kontak: "" });
  const [editId, setEditId] = useState(null);
  const [msg, setMsg] = useState("");

  const jumlahTransaksi = useMemo(() => {
    const m = {};
    for (const t of masuk) m[t.supplier] = (m[t.supplier] || 0) + 1;
    return m;
  }, [masuk]);

  const filtered = suppliers.filter((s) => !q || s.nama.toLowerCase().includes(q.toLowerCase())).sort((a, b) => a.nama.localeCompare(b.nama));

  async function submit(e) {
    e.preventDefault();
    const nama = form.nama.trim();
    if (!nama) { setMsg("Nama supplier wajib diisi."); return; }
    if (!editId && suppliers.some((s) => s.nama.toLowerCase() === nama.toLowerCase())) { setMsg("Supplier dengan nama ini sudah ada."); return; }
    try {
      await persist.simpanSupplier({ id: editId, nama, kontak: form.kontak }, editId);
      await addLog(editId ? "Ubah data supplier" : "Tambah supplier", nama);
      setForm({ nama: "", kontak: "" });
      setEditId(null);
      setMsg("Tersimpan.");
      setTimeout(() => setMsg(""), 2000);
    } catch (err) {
      setMsg("Gagal menyimpan ke Supabase: " + (err.message || String(err)));
    }
  }

  function editRow(s) {
    setEditId(s.id);
    setForm({ nama: s.nama, kontak: s.kontak || "" });
  }

  async function removeRow(s) {
    if (!confirm(`Hapus supplier "${s.nama}"? Riwayat transaksi lama tidak akan terhapus.`)) return;
    await persist.hapusSupplier(s.id);
    await addLog("Hapus supplier", s.nama);
  }

  return (
    <div>
      <PageHeader title="Data Supplier" subtitle="Daftar supplier akan otomatis bertambah saat mencatat penerimaan barang dengan nama supplier baru." />
      {msg && <div style={{ fontSize: 12.5, color: "#2F5D50", marginBottom: 14 }}>{msg}</div>}
      <div className="grid-master">
        <Card>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10 }}>{editId ? "Ubah Supplier" : "Tambah Supplier"}</div>
          <form onSubmit={submit}>
            <Field label="Nama Supplier"><input value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} style={inputStyle} placeholder="mis. CV Sumber Tani" /></Field>
            <Field label="Kontak (opsional)"><input value={form.kontak} onChange={(e) => setForm({ ...form, kontak: e.target.value })} style={inputStyle} placeholder="No. HP / alamat" /></Field>
            <div style={{ display: "flex", gap: 8 }}>
              <button type="submit" style={primaryBtn}>{editId ? <><Pencil size={15} /> Simpan Perubahan</> : <><Plus size={15} /> Tambah Supplier</>}</button>
              {editId && <button type="button" onClick={() => { setEditId(null); setForm({ nama: "", kontak: "" }); }} style={{ ...primaryBtn, background: "#EDE8DA", color: "#3A4038" }}><X size={15} /></button>}
            </div>
          </form>
        </Card>
        <Card style={{ maxHeight: 560, overflowY: "auto" }}>
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Cari supplier..." style={{ ...inputStyle, marginBottom: 10 }} />
          <div className="table-wrap">
            <table>
              <thead><tr>{["Nama Supplier", "Kontak", "Jml. Transaksi", ""].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
              <tbody>
                {filtered.map((s) => (
                  <tr key={s.id}>
                    <Td style={{ fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}><Truck size={13} color="#8B8371" /> {s.nama}</Td>
                    <Td>{s.kontak || "-"}</Td>
                    <Td className="mono">{jumlahTransaksi[s.nama] || 0}</Td>
                    <Td>
                      <div style={{ display: "flex", gap: 6 }}>
                        <button onClick={() => editRow(s)} style={iconBtn}><Pencil size={13} /></button>
                        <button onClick={() => removeRow(s)} style={{ ...iconBtn, color: "#A8402A" }}><Trash2 size={13} /></button>
                      </div>
                    </Td>
                  </tr>
                ))}
                {filtered.length === 0 && <tr><td colSpan={4}><Empty text="Belum ada data supplier." /></td></tr>}
              </tbody>
            </table>
          </div>
        </Card>
      </div>
    </div>
  );
}

/* ============================== PENGATURAN PERIODE (Super Admin) ============================== */
function PengaturanPeriode({ ctx }) {
  const { periode, computeStockFor } = ctx;
  const activeInfo = periode.list.find((p) => p.id === periode.activeId);
  const [editId, setEditId] = useState(null);
  const [editForm, setEditForm] = useState({ nama: "", mulai: "", selesai: "" });
  const [showTutup, setShowTutup] = useState(false);
  const [tutupForm, setTutupForm] = useState({ nama: "", mulai: "", selesai: "" });
  const [msg, setMsg] = useState("");

  function startEdit(p) {
    setEditId(p.id);
    setEditForm({ nama: p.nama, mulai: p.mulai, selesai: p.selesai });
  }
  async function saveEdit(e) {
    e.preventDefault();
    await periode.editPeriode(editId, editForm);
    setEditId(null);
    setMsg("Data periode diperbarui.");
    setTimeout(() => setMsg(""), 2000);
  }

  function openTutup() {
    const d = activeInfo ? new Date(activeInfo.selesai + "T00:00:00") : new Date();
    d.setDate(d.getDate() + 1);
    const mulaiBaru = d.toISOString().slice(0, 10);
    setTutupForm({ nama: `Periode ${periode.list.length + 1}`, mulai: mulaiBaru, selesai: "" });
    setShowTutup(true);
  }

  async function submitTutup(e) {
    e.preventDefault();
    if (!tutupForm.nama || !tutupForm.mulai || !tutupForm.selesai) { setMsg("Lengkapi nama dan tanggal periode baru."); return; }
    await periode.tutupPeriode(tutupForm);
    setShowTutup(false);
    setMsg("Periode baru dibuka. Saldo akhir periode sebelumnya otomatis menjadi saldo awal periode baru.");
    setTimeout(() => setMsg(""), 4000);
  }

  const sorted = [...periode.list].sort((a, b) => (a.mulai < b.mulai ? 1 : -1));

  return (
    <div>
      <PageHeader title="Pengaturan Periode" subtitle="Kelola periode pencatatan stock. Saat periode ditutup, saldo akhir otomatis menjadi saldo awal periode baru." />
      {msg && <div style={{ fontSize: 12.5, color: "#2F5D50", marginBottom: 14 }}>{msg}</div>}

      <Card style={{ marginBottom: 16 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10, marginBottom: 4 }}>
          <div>
            <div className="sg" style={{ fontWeight: 700, fontSize: 14.5 }}>Periode Aktif Saat Ini</div>
            <div style={{ fontSize: 13, color: "#5C5646", marginTop: 4 }}>{activeInfo ? `${activeInfo.nama} — ${fmtTgl(activeInfo.mulai)} s.d. ${fmtTgl(activeInfo.selesai)}` : "-"}</div>
          </div>
          <button onClick={openTutup} style={{ ...primaryBtn, width: "auto", padding: "9px 16px", background: "#C1602E" }}>
            <Lock size={14} /> Tutup Periode &amp; Buka Baru
          </button>
        </div>
      </Card>

      {showTutup && (
        <Card style={{ marginBottom: 16, border: "1px solid #C1602E" }}>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10 }}>Tutup Periode &amp; Buka Periode Baru</div>
          <p style={{ fontSize: 12.5, color: "#7A7362", marginTop: -4, marginBottom: 12 }}>
            Saldo akhir setiap barang pada periode aktif akan dihitung otomatis dan dijadikan saldo awal untuk periode baru ini.
          </p>
          <form onSubmit={submitTutup}>
            <div className="grid-eq2">
              <Field label="Nama Periode Baru"><input value={tutupForm.nama} onChange={(e) => setTutupForm({ ...tutupForm, nama: e.target.value })} style={inputStyle} /></Field>
              <div />
              <Field label="Tanggal Mulai"><input type="date" value={tutupForm.mulai} onChange={(e) => setTutupForm({ ...tutupForm, mulai: e.target.value })} style={inputStyle} /></Field>
              <Field label="Tanggal Selesai"><input type="date" value={tutupForm.selesai} onChange={(e) => setTutupForm({ ...tutupForm, selesai: e.target.value })} style={inputStyle} /></Field>
            </div>
            <div style={{ display: "flex", gap: 8 }}>
              <button type="submit" style={{ ...primaryBtn, width: "auto", padding: "9px 16px" }}>Konfirmasi &amp; Buka Periode Baru</button>
              <button type="button" onClick={() => setShowTutup(false)} style={{ ...primaryBtn, width: "auto", padding: "9px 16px", background: "#EDE8DA", color: "#3A4038" }}>Batal</button>
            </div>
          </form>
        </Card>
      )}

      <Card>
        <div className="sg" style={{ fontWeight: 700, marginBottom: 10 }}>Riwayat Periode</div>
        <div className="table-wrap">
          <table>
            <thead><tr>{["Nama Periode", "Mulai", "Selesai", "Status", "Total Nilai Akhir", ""].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
            <tbody>
              {sorted.map((p) => {
                const nilaiAkhir = computeStockFor(p.id).reduce((s, i) => s + i.nilai, 0);
                const editing = editId === p.id;
                return (
                  <tr key={p.id}>
                    {editing ? (
                      <>
                        <Td colSpan={3}>
                          <form onSubmit={saveEdit} style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                            <input value={editForm.nama} onChange={(e) => setEditForm({ ...editForm, nama: e.target.value })} style={{ ...inputStyle, width: 160 }} />
                            <input type="date" value={editForm.mulai} onChange={(e) => setEditForm({ ...editForm, mulai: e.target.value })} style={{ ...inputStyle, width: 140 }} />
                            <input type="date" value={editForm.selesai} onChange={(e) => setEditForm({ ...editForm, selesai: e.target.value })} style={{ ...inputStyle, width: 140 }} />
                            <button type="submit" style={{ ...iconBtn, background: "#2F5D50", color: "#fff" }}>Simpan</button>
                            <button type="button" onClick={() => setEditId(null)} style={iconBtn}><X size={13} /></button>
                          </form>
                        </Td>
                      </>
                    ) : (
                      <>
                        <Td style={{ fontWeight: 600 }}>{p.nama}</Td>
                        <Td>{fmtTgl(p.mulai)}</Td>
                        <Td>{fmtTgl(p.selesai)}</Td>
                      </>
                    )}
                    {!editing && (
                      <>
                        <Td>{p.id === periode.activeId ? <span style={{ fontSize: 10.5, fontWeight: 700, color: "#1B2E27", background: "#6FA88F", padding: "3px 8px", borderRadius: 5 }}>AKTIF</span> : <span style={{ fontSize: 10.5, fontWeight: 700, color: "#fff", background: "#8B8371", padding: "3px 8px", borderRadius: 5 }}>ARSIP</span>}</Td>
                        <Td className="mono">{fmtRp(nilaiAkhir)}</Td>
                        <Td><button onClick={() => startEdit(p)} style={iconBtn}><Pencil size={13} /></button></Td>
                      </>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

/* ============================== MANAJEMEN USER (Super Admin) ============================== */
function ManajemenUser({ ctx }) {
  const { session, addLog } = ctx;
  const isSuperAdmin = session.role === "superadmin";
  const [profiles, setProfiles] = useState([]);
  const [loadingProfiles, setLoadingProfiles] = useState(true);
  const [form, setForm] = useState({ nama: "", username: "", password: "", role: "admin" });
  const [pwForm, setPwForm] = useState({ password: "" });
  const [pinForm, setPinForm] = useState("");
  const [msg, setMsg] = useState("");
  const [busy, setBusy] = useState(false);

  const loadProfiles = useCallback(async () => {
    setLoadingProfiles(true);
    const { data } = await supabase.from("profiles").select("*").order("created_at", { ascending: true });
    setProfiles(data || []);
    setLoadingProfiles(false);
  }, []);

  useEffect(() => { if (isSuperAdmin) loadProfiles(); }, [isSuperAdmin, loadProfiles]);

  async function addProfile(e) {
    e.preventDefault();
    if (!form.nama || !form.username || !form.password) { setMsg("Lengkapi nama, username, dan password."); return; }
    setBusy(true);
    const { data, error } = await supabase.functions.invoke("create-user", { body: form });
    setBusy(false);
    if (error || data?.error) { setMsg(data?.error || error.message); return; }
    await addLog("Tambah akun", `${form.nama} (${form.username}, ${form.role})`);
    setForm({ nama: "", username: "", password: "", role: "admin" });
    setMsg("Akun ditambahkan.");
    setTimeout(() => setMsg(""), 2500);
    loadProfiles();
  }

  async function removeProfile(id, nama) {
    if (!confirm(`Hapus akun "${nama}"? Login-nya juga akan dihapus sepenuhnya.`)) return;
    setBusy(true);
    const { data, error } = await supabase.functions.invoke("delete-user", { body: { id } });
    setBusy(false);
    if (error || data?.error) { setMsg(data?.error || error.message); return; }
    await addLog("Hapus akun", nama);
    loadProfiles();
  }

  async function changeMyPassword(e) {
    e.preventDefault();
    if (!pwForm.password || pwForm.password.length < 6) { setMsg("Password baru minimal 6 karakter."); return; }
    const { error } = await supabase.auth.updateUser({ password: pwForm.password });
    if (error) { setMsg(error.message); return; }
    await addLog("Ubah password sendiri", session.username);
    setPwForm({ password: "" });
    setMsg("Password berhasil diubah.");
    setTimeout(() => setMsg(""), 2500);
  }

  async function savePin(e) {
    e.preventDefault();
    if (!pinForm) { setMsg("Isi PIN baru."); return; }
    const { error } = await supabase.rpc("set_viewer_pin", { p_new_pin: pinForm });
    if (error) { setMsg(error.message); return; }
    await addLog("Ubah PIN viewer", "***");
    setPinForm("");
    setMsg("PIN viewer diperbarui.");
    setTimeout(() => setMsg(""), 2500);
  }

  return (
    <div>
      <PageHeader title="Manajemen User" subtitle="Kelola akun Admin/Super Admin dan PIN Viewer." />
      {msg && <div style={{ fontSize: 12.5, color: "#2F5D50", marginBottom: 14 }}>{msg}</div>}
      <div className="grid-eq2">
        <div style={{ display: "grid", gap: 16 }}>
          {isSuperAdmin && (
            <Card>
              <div className="sg" style={{ fontWeight: 700, marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}><Users size={15} /> Akun Admin / Super Admin</div>
              <div style={{ fontSize: 11.5, color: "#8B8371", lineHeight: 1.5, marginBottom: 10 }}>
                Akun login lengkap (username + password) dibuat langsung dari sini, tanpa perlu email.
              </div>
              <form onSubmit={addProfile} style={{ marginBottom: 14, display: "grid", gap: 8 }}>
                <input value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} placeholder="Nama lengkap" style={inputStyle} />
                <input value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} placeholder="Username" style={inputStyle} />
                <input type="password" value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} placeholder="Password (min. 6 karakter)" style={inputStyle} />
                <select value={form.role} onChange={(e) => setForm({ ...form, role: e.target.value })} style={inputStyle}>
                  <option value="admin">Admin</option>
                  <option value="superadmin">Super Admin</option>
                </select>
                <button type="submit" disabled={busy} style={{ ...primaryBtn, width: "auto", padding: "9px 16px" }}><Plus size={15} /> {busy ? "Memproses..." : "Tambah Akun"}</button>
              </form>
              {loadingProfiles ? <Empty text="Memuat..." /> : profiles.map((p) => (
                <div key={p.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #EEE9DB" }}>
                  <div>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{p.nama} <RoleBadge role={p.role} /></div>
                    <div style={{ fontSize: 11.5, color: "#8B8371" }} className="mono">{p.username}</div>
                  </div>
                  <button onClick={() => removeProfile(p.id, p.nama)} disabled={busy} style={{ ...iconBtn, color: "#A8402A" }}><Trash2 size={13} /></button>
                </div>
              ))}
              {!loadingProfiles && profiles.length === 0 && <Empty text="Belum ada akun terdaftar." />}
            </Card>
          )}

          <Card>
            <div className="sg" style={{ fontWeight: 700, marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}><ShieldCheck size={15} /> Ganti Password Saya</div>
            <form onSubmit={changeMyPassword} style={{ display: "grid", gap: 8 }}>
              <Field label="Password baru"><input type="password" value={pwForm.password} onChange={(e) => setPwForm({ password: e.target.value })} style={inputStyle} /></Field>
              <button type="submit" style={{ ...primaryBtn, width: "auto", padding: "9px 16px" }}>Simpan</button>
            </form>
          </Card>
        </div>

        <div style={{ display: "grid", gap: 16 }}>
          {isSuperAdmin && (
            <Card>
              <div className="sg" style={{ fontWeight: 700, marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}><KeyRound size={15} /> PIN Viewer</div>
              <form onSubmit={savePin} style={{ display: "grid", gap: 8 }}>
                <Field label="PIN baru"><input value={pinForm} onChange={(e) => setPinForm(e.target.value.replace(/\D/g, ""))} inputMode="numeric" style={inputStyle} placeholder="mis. 1234" /></Field>
                <button type="submit" style={{ ...primaryBtn, width: "auto", padding: "9px 16px" }}>Simpan PIN</button>
              </form>
            </Card>
          )}
          <div style={{ fontSize: 11.5, color: "#8B8371", lineHeight: 1.5, padding: "0 4px" }}>
            Kredensial staff dikelola oleh Supabase Auth dan tidak pernah tersimpan di database aplikasi sebagai teks biasa.
          </div>
        </div>
      </div>
    </div>
  );
}

/* ============================== LOG AKTIVITAS (Super Admin) ============================== */
function LogAktivitas({ ctx }) {
  const { log } = ctx;
  return (
    <div>
      <PageHeader title="Log Aktivitas" subtitle="Riwayat seluruh aksi admin dan super admin beserta waktu kejadian." />
      <Card><div className="table-wrap">
        <table>
          <thead><tr>{["Waktu", "Oleh", "Peran", "Aksi", "Detail"].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
          <tbody>
            {log.map((l) => (
              <tr key={l.id}>
                <Td className="mono" style={{ fontSize: 11.5 }}>{fmtDate(l.waktu)}</Td>
                <Td>{l.oleh}</Td>
                <Td><RoleBadge role={l.role} /></Td>
                <Td>{l.aksi}</Td>
                <Td style={{ color: "#7A7362" }}>{l.detail}</Td>
              </tr>
            ))}
            {log.length === 0 && <tr><td colSpan={5}><Empty text="Belum ada aktivitas tercatat." /></td></tr>}
          </tbody>
        </table>
      </div></Card>
    </div>
  );
}

function Th({ children }) { return <th style={{ textAlign: "left", fontSize: 11, color: "#8B8371", textTransform: "uppercase", letterSpacing: 0.3, padding: "8px 10px", borderBottom: "2px solid #E4DFCF" }}>{children}</th>; }
function Td({ children, className, style, colSpan }) { return <td className={className} colSpan={colSpan} style={{ padding: "9px 10px", fontSize: 13, borderBottom: "1px solid #EEE9DB", ...style }}>{children}</td>; }
