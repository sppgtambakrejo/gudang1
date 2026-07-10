import React, { useState, useEffect, useMemo, useCallback } from "react";
import { Package, LogIn, LogOut, LayoutDashboard, ArrowDownToLine, ArrowUpFromLine, ClipboardList, Users, ScrollText, Search, Plus, Trash2, Pencil, X, KeyRound, ShieldCheck, Eye, AlertTriangle, Menu, CalendarRange, Download, Lock, ChevronDown, Truck } from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";

/* ============================== MASTER DATA (seeded from SPPG Sleman Tempel Tambakrejo, Periode 16-28 Feb 2026) ============================== */
const ITEMS = [{"kode": "KH.01.001", "nama": "Beras putih (premium)", "satuan": "kg"}, {"kode": "KH.01.002", "nama": "Beras putih (medium)", "satuan": "kg"}, {"kode": "KH.01.003", "nama": "Beras merah", "satuan": "kg"}, {"kode": "KH.01.004", "nama": "Beras hitam", "satuan": "kg"}, {"kode": "KH.01.005", "nama": "Beras ketan putih", "satuan": "kg"}, {"kode": "KH.01.006", "nama": "Beras ketan hitam", "satuan": "kg"}, {"kode": "KH.01.007", "nama": "Tepung beras", "satuan": "kg"}, {"kode": "KH.01.008", "nama": "Tepung ketan", "satuan": "kg"}, {"kode": "KH.01.009", "nama": "Mie beras", "satuan": "kg"}, {"kode": "KH.01.099", "nama": "Beras dan olahan padi-lainnya", "satuan": "kg"}, {"kode": "KH.02.001", "nama": "Jagung pipil kering", "satuan": "kg"}, {"kode": "KH.02.002", "nama": "Jagung pipil segar", "satuan": "kg"}, {"kode": "KH.02.003", "nama": "Jagung manis segar", "satuan": "kg"}, {"kode": "KH.02.004", "nama": "Beras jagung", "satuan": "kg"}, {"kode": "KH.02.005", "nama": "Jagung giling", "satuan": "kg"}, {"kode": "KH.02.006", "nama": "Tepung jagung (maizena)", "satuan": "kg"}, {"kode": "KH.02.007", "nama": "Tepung jagung lokal (cornmeal)", "satuan": "kg"}, {"kode": "KH.02.099", "nama": "Jagung dan olahan-lainnya", "satuan": "kg"}, {"kode": "KH.03.001", "nama": "Singkong segar", "satuan": "kg"}, {"kode": "KH.03.002", "nama": "Singkong kering/gaplek", "satuan": "kg"}, {"kode": "KH.03.003", "nama": "Tepung singkong (tapioka)", "satuan": "kg"}, {"kode": "KH.03.004", "nama": "Mocaf", "satuan": "kg"}, {"kode": "KH.03.005", "nama": "Ubi jalar putih", "satuan": "kg"}, {"kode": "KH.03.006", "nama": "Ubi jalar kuning/oranye", "satuan": "kg"}, {"kode": "KH.03.007", "nama": "Ubi jalar ungu", "satuan": "kg"}, {"kode": "KH.03.008", "nama": "Kentang", "satuan": "kg"}, {"kode": "KH.03.009", "nama": "Talas", "satuan": "kg"}, {"kode": "KH.03.010", "nama": "Ganyong", "satuan": "kg"}, {"kode": "KH.03.011", "nama": "Sagu basah", "satuan": "kg"}, {"kode": "KH.03.012", "nama": "Tepung sagu", "satuan": "kg"}, {"kode": "KH.03.099", "nama": "Umbi umbian-lainnya", "satuan": "kg"}, {"kode": "KH.04.001", "nama": "Tepung terigu protein rendah", "satuan": "kg"}, {"kode": "KH.04.002", "nama": "Tepung terigu protein sedang", "satuan": "kg"}, {"kode": "KH.04.003", "nama": "Tepung terigu protein tinggi", "satuan": "kg"}, {"kode": "KH.04.004", "nama": "Tepung semolina", "satuan": "kg"}, {"kode": "KH.04.005", "nama": "Mie kering (non-instan)", "satuan": "pak"}, {"kode": "KH.04.006", "nama": "Spageti", "satuan": "pak"}, {"kode": "KH.04.007", "nama": "Makaroni", "satuan": "pak"}, {"kode": "KH.04.008", "nama": "Roti tawar", "satuan": "pak"}, {"kode": "KH.04.009", "nama": "Roti gandum", "satuan": "pak"}, {"kode": "KH.04.010", "nama": "Tepung panir", "satuan": "pak"}, {"kode": "KH.04.011", "nama": "Biskuit", "satuan": "pak"}, {"kode": "KH.04.012", "nama": "Bubur bayi", "satuan": "pak"}, {"kode": "KH.04.099", "nama": "Gandum dan olahan tepung-lainnya", "satuan": "pak"}, {"kode": "KH.05.001", "nama": "Oat (rolled oats)", "satuan": "kg"}, {"kode": "KH.05.002", "nama": "Oat instan", "satuan": "kg"}, {"kode": "KH.05.003", "nama": "Sorgum biji", "satuan": "kg"}, {"kode": "KH.05.004", "nama": "Tepung sorgum", "satuan": "kg"}, {"kode": "KH.05.005", "nama": "Millet/jewawut biji", "satuan": "kg"}, {"kode": "KH.05.006", "nama": "Tepung millet", "satuan": "kg"}, {"kode": "KH.05.099", "nama": "Sereal-lainnya", "satuan": "kg"}, {"kode": "PH.01.001", "nama": "Ayam broiler utuh", "satuan": "kg"}, {"kode": "PH.01.002", "nama": "Fillet dada ayam", "satuan": "kg"}, {"kode": "PH.01.003", "nama": "Paha ayam atas", "satuan": "kg"}, {"kode": "PH.01.004", "nama": "Paha ayam bawah (drumstick)", "satuan": "kg"}, {"kode": "PH.01.005", "nama": "Sayap ayam", "satuan": "kg"}, {"kode": "PH.01.006", "nama": "Hati ayam", "satuan": "kg"}, {"kode": "PH.01.007", "nama": "Ampela ayam", "satuan": "kg"}, {"kode": "PH.01.008", "nama": "Telur ayam ras/negeri", "satuan": "kg"}, {"kode": "PH.01.009", "nama": "Telur ayam kampung", "satuan": "kg"}, {"kode": "PH.01.010", "nama": "Telur puyuh", "satuan": "kg"}, {"kode": "PH.01.011", "nama": "Sosis ayam", "satuan": "pak"}, {"kode": "PH.01.099", "nama": "Ayam dan olahannya-lainnya", "satuan": "kg"}, {"kode": "PH.02.001", "nama": "Daging sapi segar (has luar/sirloin)", "satuan": "kg"}, {"kode": "PH.02.002", "nama": "Daging sapi segar (has dalam/tenderloin)", "satuan": "kg"}, {"kode": "PH.02.003", "nama": "Daging sapi segar (sandung lamur/brisket)", "satuan": "kg"}, {"kode": "PH.02.004", "nama": "Daging sapi segar (paha)", "satuan": "kg"}, {"kode": "PH.02.005", "nama": "Daging sapi giling", "satuan": "kg"}, {"kode": "PH.02.006", "nama": "Hati sapi", "satuan": "kg"}, {"kode": "PH.02.007", "nama": "Iga sapi", "satuan": "kg"}, {"kode": "PH.02.008", "nama": "Daging kerbau segar", "satuan": "kg"}, {"kode": "PH.02.009", "nama": "Daging kambing segar", "satuan": "kg"}, {"kode": "PH.02.010", "nama": "Daging domba segar", "satuan": "kg"}, {"kode": "PH.02.011", "nama": "Hati kambing", "satuan": "kg"}, {"kode": "PH.02.012", "nama": "Sosis sapi", "satuan": "pak"}, {"kode": "PH.02.099", "nama": "Daging sapi/kerbau/kambing-lainnya", "satuan": "kg"}, {"kode": "PH.03.001", "nama": "Lele", "satuan": "kg"}, {"kode": "PH.03.002", "nama": "Nila", "satuan": "kg"}, {"kode": "PH.03.003", "nama": "Mujair", "satuan": "kg"}, {"kode": "PH.03.004", "nama": "Patin", "satuan": "kg"}, {"kode": "PH.03.005", "nama": "Gurame", "satuan": "kg"}, {"kode": "PH.03.006", "nama": "Mas", "satuan": "kg"}, {"kode": "PH.03.007", "nama": "Gabus", "satuan": "kg"}, {"kode": "PH.03.008", "nama": "Fillet ikan tawar", "satuan": "kg"}, {"kode": "PH.03.099", "nama": "Ikan air tawar-lainnya", "satuan": "kg"}, {"kode": "PH.04.001", "nama": "Kembung", "satuan": "kg"}, {"kode": "PH.04.002", "nama": "Tongkol", "satuan": "kg"}, {"kode": "PH.04.003", "nama": "Cakalang", "satuan": "kg"}, {"kode": "PH.04.004", "nama": "Tuna", "satuan": "kg"}, {"kode": "PH.04.005", "nama": "Sarden/lemuru", "satuan": "kg"}, {"kode": "PH.04.006", "nama": "Bandeng", "satuan": "kg"}, {"kode": "PH.04.007", "nama": "Tenggiri", "satuan": "kg"}, {"kode": "PH.04.008", "nama": "Kakap", "satuan": "kg"}, {"kode": "PH.04.009", "nama": "Fillet ikan laut", "satuan": "kg"}, {"kode": "PH.04.010", "nama": "Ikan laut-lainnya", "satuan": "kg"}, {"kode": "PH.05.001", "nama": "Udang segar", "satuan": "kg"}, {"kode": "PH.05.002", "nama": "Udang beku", "satuan": "kg"}, {"kode": "PH.05.003", "nama": "Cumi-cumi", "satuan": "kg"}, {"kode": "PH.05.004", "nama": "Sotong", "satuan": "kg"}, {"kode": "PH.05.005", "nama": "Kepiting/rajungan", "satuan": "kg"}, {"kode": "PH.05.006", "nama": "Kerang hijau", "satuan": "kg"}, {"kode": "PH.05.007", "nama": "Kerang dara", "satuan": "kg"}, {"kode": "PH.05.099", "nama": "Produk perikanan lain-lainnya", "satuan": "kg"}, {"kode": "PH.06.001", "nama": "Susu UHT full cream", "satuan": "liter"}, {"kode": "PH.06.002", "nama": "Susu UHT rendah lemak", "satuan": "liter"}, {"kode": "PH.06.003", "nama": "Susu pasteurisasi", "satuan": "liter"}, {"kode": "PH.06.004", "nama": "Susu bubuk", "satuan": "kg"}, {"kode": "PH.06.005", "nama": "Yogurt plain", "satuan": "pak"}, {"kode": "PH.06.006", "nama": "Keju cheddar", "satuan": "pak"}, {"kode": "PH.06.007", "nama": "Keju olahan (slice)", "satuan": "pak"}, {"kode": "PH.06.099", "nama": "Susu dan olahan-lainnya", "satuan": "liter"}, {"kode": "PN.01.001", "nama": "Kedelai biji kering", "satuan": "kg"}, {"kode": "PN.01.002", "nama": "Tempe", "satuan": "kg"}, {"kode": "PN.01.003", "nama": "Tahu putih", "satuan": "kg"}, {"kode": "PN.01.004", "nama": "Tahu kuning", "satuan": "kg"}, {"kode": "PN.01.005", "nama": "Tahu sutra", "satuan": "kg"}, {"kode": "PN.01.006", "nama": "Susu kedelai", "satuan": "liter"}, {"kode": "PN.01.007", "nama": "Tepung kedelai", "satuan": "kg"}, {"kode": "PN.01.099", "nama": "Kedelai dan olahannya-lainnya", "satuan": "kg"}, {"kode": "PN.02.001", "nama": "Kacang tanah kupas", "satuan": "kg"}, {"kode": "PN.02.002", "nama": "Kacang tanah dengan kulit", "satuan": "kg"}, {"kode": "PN.02.003", "nama": "Kacang hijau", "satuan": "kg"}, {"kode": "PN.02.004", "nama": "Kacang merah", "satuan": "kg"}, {"kode": "PN.02.005", "nama": "Kacang tolo/tunggak", "satuan": "kg"}, {"kode": "PN.02.006", "nama": "Kacang hitam", "satuan": "kg"}, {"kode": "PN.02.007", "nama": "Kacang arab/chickpea", "satuan": "kg"}, {"kode": "PN.02.008", "nama": "Kacang polong/pea", "satuan": "kg"}, {"kode": "PN.02.009", "nama": "Kacang mede", "satuan": "kg"}, {"kode": "PN.02.010", "nama": "Kacang almond", "satuan": "kg"}, {"kode": "PN.02.099", "nama": "Kacang kacangan-lainnya", "satuan": "kg"}, {"kode": "PN.03.001", "nama": "Wijen", "satuan": "kg"}, {"kode": "PN.03.002", "nama": "Biji bunga matahari", "satuan": "kg"}, {"kode": "PN.03.003", "nama": "Chia seed", "satuan": "kg"}, {"kode": "PN.03.004", "nama": "Flaxseed", "satuan": "kg"}, {"kode": "PN.03.099", "nama": "Biji-bijian dan produknya-lainnya", "satuan": "kg"}, {"kode": "SY.01.001", "nama": "Bayam", "satuan": "ikat"}, {"kode": "SY.01.002", "nama": "Kangkung", "satuan": "ikat"}, {"kode": "SY.01.003", "nama": "Sawi hijau", "satuan": "ikat"}, {"kode": "SY.01.004", "nama": "Sawi putih", "satuan": "ikat"}, {"kode": "SY.01.005", "nama": "Pakcoy", "satuan": "ikat"}, {"kode": "SY.01.006", "nama": "Selada", "satuan": "ikat"}, {"kode": "SY.01.007", "nama": "Daun singkong", "satuan": "ikat"}, {"kode": "SY.01.008", "nama": "Daun pepaya", "satuan": "ikat"}, {"kode": "SY.01.009", "nama": "Kelor", "satuan": "ikat"}, {"kode": "SY.01.010", "nama": "Katuk", "satuan": "ikat"}, {"kode": "SY.01.011", "nama": "Seledri daun", "satuan": "ikat"}, {"kode": "SY.01.012", "nama": "Brokoli", "satuan": "kg"}, {"kode": "SY.01.013", "nama": "Kembang kol", "satuan": "kg"}, {"kode": "SY.01.014", "nama": "Buncis", "satuan": "kg"}, {"kode": "SY.01.015", "nama": "Seledri batang", "satuan": "ikat"}, {"kode": "SY.01.016", "nama": "Asparagus", "satuan": "ikat"}, {"kode": "SY.01.099", "nama": "Sayuran daun/batang/bunga-lainnya", "satuan": "ikat"}, {"kode": "SY.02.001", "nama": "Tomat", "satuan": "kg"}, {"kode": "SY.02.002", "nama": "Mentimun", "satuan": "kg"}, {"kode": "SY.02.003", "nama": "Terong ungu", "satuan": "kg"}, {"kode": "SY.02.004", "nama": "Terong hijau", "satuan": "kg"}, {"kode": "SY.02.005", "nama": "Cabai merah besar", "satuan": "kg"}, {"kode": "SY.02.006", "nama": "Cabai rawit", "satuan": "kg"}, {"kode": "SY.02.007", "nama": "Labu siam", "satuan": "kg"}, {"kode": "SY.02.008", "nama": "Pare", "satuan": "kg"}, {"kode": "SY.02.009", "nama": "Paprika", "satuan": "kg"}, {"kode": "SY.02.010", "nama": "Okra", "satuan": "kg"}, {"kode": "SY.02.011", "nama": "Wortel", "satuan": "kg"}, {"kode": "SY.02.012", "nama": "Lobak putih", "satuan": "kg"}, {"kode": "SY.02.013", "nama": "Bit", "satuan": "kg"}, {"kode": "SY.02.099", "nama": "Sayuran buah/umbi/akar-lainnya", "satuan": "kg"}, {"kode": "SY.03.001", "nama": "Bawang merah", "satuan": "kg"}, {"kode": "SY.03.002", "nama": "Bawang putih", "satuan": "kg"}, {"kode": "SY.03.003", "nama": "Bawang bombai", "satuan": "kg"}, {"kode": "SY.03.004", "nama": "Daun bawang", "satuan": "kg"}, {"kode": "SY.03.005", "nama": "Bawang prei", "satuan": "kg"}, {"kode": "SY.03.006", "nama": "Jahe", "satuan": "kg"}, {"kode": "SY.03.007", "nama": "Kunyit", "satuan": "kg"}, {"kode": "SY.03.008", "nama": "Lengkuas", "satuan": "kg"}, {"kode": "SY.03.009", "nama": "Kencur", "satuan": "kg"}, {"kode": "SY.03.010", "nama": "Serai", "satuan": "kg"}, {"kode": "SY.03.011", "nama": "Temu kunci", "satuan": "kg"}, {"kode": "SY.03.012", "nama": "Kemiri", "satuan": "kg"}, {"kode": "SY.03.099", "nama": "Sayuran bawang dan aromatik lainnya", "satuan": "kg"}, {"kode": "SY.04.001", "nama": "Jamur tiram", "satuan": "kg"}, {"kode": "SY.04.002", "nama": "Jamur kancing", "satuan": "kg"}, {"kode": "SY.04.003", "nama": "Jamur kuping", "satuan": "kg"}, {"kode": "SY.04.004", "nama": "Jamur shiitake", "satuan": "kg"}, {"kode": "SY.04.099", "nama": "Jamur-lainnya", "satuan": "kg"}, {"kode": "BU.01.001", "nama": "Pisang ambon", "satuan": "kg"}, {"kode": "BU.01.002", "nama": "Pisang kepok", "satuan": "kg"}, {"kode": "BU.01.003", "nama": "Pisang raja", "satuan": "kg"}, {"kode": "BU.01.004", "nama": "Pepaya", "satuan": "kg"}, {"kode": "BU.01.005", "nama": "Semangka", "satuan": "kg"}, {"kode": "BU.01.006", "nama": "Melon", "satuan": "kg"}, {"kode": "BU.01.007", "nama": "Mangga harum manis", "satuan": "kg"}, {"kode": "BU.01.008", "nama": "Mangga gedong", "satuan": "kg"}, {"kode": "BU.01.009", "nama": "Jeruk manis", "satuan": "kg"}, {"kode": "BU.01.010", "nama": "Jeruk keprok", "satuan": "kg"}, {"kode": "BU.01.011", "nama": "Jambu biji", "satuan": "kg"}, {"kode": "BU.01.012", "nama": "Jambu air", "satuan": "kg"}, {"kode": "BU.01.013", "nama": "Nanas", "satuan": "kg"}, {"kode": "BU.01.014", "nama": "Salak", "satuan": "kg"}, {"kode": "BU.01.015", "nama": "Rambutan", "satuan": "kg"}, {"kode": "BU.01.016", "nama": "Duku", "satuan": "kg"}, {"kode": "BU.01.017", "nama": "Sirsak", "satuan": "kg"}, {"kode": "BU.01.018", "nama": "Durian", "satuan": "kg"}, {"kode": "BU.01.019", "nama": "Manggis", "satuan": "kg"}, {"kode": "BU.01.020", "nama": "Alpukat", "satuan": "kg"}, {"kode": "BU.01.021", "nama": "Apel malang", "satuan": "kg"}, {"kode": "BU.01.022", "nama": "Buah naga merah", "satuan": "kg"}, {"kode": "BU.01.023", "nama": "Buah naga putih", "satuan": "kg"}, {"kode": "BU.01.024", "nama": "Kedondong", "satuan": "kg"}, {"kode": "BU.01.025", "nama": "Markisa", "satuan": "kg"}, {"kode": "BU.01.026", "nama": "Belimbing", "satuan": "kg"}, {"kode": "BU.01.027", "nama": "Sawo", "satuan": "kg"}, {"kode": "BU.01.028", "nama": "Matoa", "satuan": "kg"}, {"kode": "BU.01.099", "nama": "Buah lokal-lainnya", "satuan": "kg"}, {"kode": "BU.02.001", "nama": "Apel", "satuan": "kg"}, {"kode": "BU.02.002", "nama": "Pear", "satuan": "kg"}, {"kode": "BU.02.003", "nama": "Anggur merah", "satuan": "kg"}, {"kode": "BU.02.004", "nama": "Anggur hijau", "satuan": "kg"}, {"kode": "BU.02.005", "nama": "Kiwi", "satuan": "kg"}, {"kode": "BU.02.099", "nama": "Buah impor-lainnya", "satuan": "kg"}, {"kode": "BB.01.001", "nama": "Minyak goreng sawit", "satuan": "liter"}, {"kode": "BB.01.002", "nama": "Minyak goreng campuran", "satuan": "liter"}, {"kode": "BB.01.003", "nama": "Minyak kelapa", "satuan": "liter"}, {"kode": "BB.01.004", "nama": "Margarin", "satuan": "kg"}, {"kode": "BB.01.005", "nama": "Mentega (butter)", "satuan": "kg"}, {"kode": "BB.01.099", "nama": "Minyak dan lemak-lainnya", "satuan": "kg"}, {"kode": "BB.02.001", "nama": "Garam beryodium", "satuan": "pak"}, {"kode": "BB.02.002", "nama": "Gula pasir", "satuan": "kg"}, {"kode": "BB.02.003", "nama": "Gula merah/aren", "satuan": "kg"}, {"kode": "BB.02.004", "nama": "Madu", "satuan": "kg"}, {"kode": "BB.02.099", "nama": "Garam, gula, dan pemanis-lainnya", "satuan": "kg"}, {"kode": "BB.03.001", "nama": "Lada/merica bubuk", "satuan": "pcs"}, {"kode": "BB.03.002", "nama": "Ketumbar bubuk", "satuan": "pcs"}, {"kode": "BB.03.003", "nama": "Jinten bubuk", "satuan": "pcs"}, {"kode": "BB.03.004", "nama": "Pala bubuk", "satuan": "pcs"}, {"kode": "BB.03.005", "nama": "Kayu manis", "satuan": "pcs"}, {"kode": "BB.03.006", "nama": "Cengkeh", "satuan": "pcs"}, {"kode": "BB.03.007", "nama": "Kapulaga", "satuan": "pcs"}, {"kode": "BB.03.008", "nama": "Daun salam kering", "satuan": "pcs"}, {"kode": "BB.03.009", "nama": "Oregano", "satuan": "pcs"}, {"kode": "BB.03.099", "nama": "Bumbu kering (rempah)-lainnya", "satuan": "pcs"}, {"kode": "BB.04.001", "nama": "Kecap manis", "satuan": "botol"}, {"kode": "BB.04.002", "nama": "Kecap asin", "satuan": "botol"}, {"kode": "BB.04.003", "nama": "Saus sambal", "satuan": "botol"}, {"kode": "BB.04.004", "nama": "Saus tomat", "satuan": "botol"}, {"kode": "BB.04.005", "nama": "Saus tiram", "satuan": "botol"}, {"kode": "BB.04.006", "nama": "Cuka makan", "satuan": "botol"}, {"kode": "BB.04.007", "nama": "Mayones", "satuan": "botol"}, {"kode": "BB.04.099", "nama": "Bumbu cair, saus, dan kondimen-lainnya", "satuan": "botol"}, {"kode": "BB.05.001", "nama": "Santan segar", "satuan": "pcs"}, {"kode": "BB.05.002", "nama": "Santan instan cair", "satuan": "pcs"}, {"kode": "BB.05.003", "nama": "Santan bubuk", "satuan": "pcs"}, {"kode": "BB.05.004", "nama": "Kaldu bubuk ayam", "satuan": "pcs"}, {"kode": "BB.05.005", "nama": "Kaldu bubuk sapi", "satuan": "pcs"}, {"kode": "BB.05.006", "nama": "Kaldu jamur", "satuan": "pcs"}, {"kode": "BB.05.007", "nama": "Agar-agar bubuk", "satuan": "pcs"}, {"kode": "BB.05.008", "nama": "Gelatin", "satuan": "pcs"}, {"kode": "BB.05.099", "nama": "Produk olahan dasar-lainnya", "satuan": "pcs"}, {"kode": "BB.06.001", "nama": "Terasi", "satuan": "pcs"}, {"kode": "BB.06.002", "nama": "Tauco", "satuan": "pcs"}, {"kode": "BB.06.003", "nama": "Ragi roti", "satuan": "pcs"}, {"kode": "BB.06.004", "nama": "Baking powder", "satuan": "pcs"}, {"kode": "BB.06.099", "nama": "Bahan fermentasi/pelengkap-lainnya", "satuan": "pcs"}, {"kode": "BB.07.001", "nama": "Air minum kemasan galon", "satuan": "pcs"}, {"kode": "BB.07.002", "nama": "Air minum kemasan botol 2 liter", "satuan": "dus"}, {"kode": "BB.07.003", "nama": "Air minum kemasan botol 1,5 liter", "satuan": "dus"}, {"kode": "BB.07.004", "nama": "Air minum kemasan botol 600 ml", "satuan": "dus"}, {"kode": "BB.07.005", "nama": "Air minum kemasan botol 330 ml", "satuan": "dus"}, {"kode": "BB.07.006", "nama": "Air minum kemasan cup", "satuan": "dus"}, {"kode": "BB.07.007", "nama": "Minuman kemasan botol/kaleng", "satuan": "pcs"}, {"kode": "BB.07.008", "nama": "Susu kental manis", "satuan": "pcs"}, {"kode": "BB.07.099", "nama": "Bahan minuman pendamping-lainnya", "satuan": "pcs"}];
const SALDO_AWAL = {"KH.01.001": {"saldo": 12, "harga": 15000}, "KH.02.004": {"saldo": 8, "harga": 17000}, "KH.03.008": {"saldo": 5, "harga": 7000}, "KH.04.001": {"saldo": 7, "harga": 14000}, "KH.04.009": {"saldo": 15, "harga": 15000}, "KH.04.011": {"saldo": 5, "harga": 7500}, "KH.04.012": {"saldo": 4, "harga": 12500}, "PH.01.001": {"saldo": 5, "harga": 30000}, "PH.01.008": {"saldo": 10, "harga": 32000}, "PH.01.010": {"saldo": 5, "harga": 22000}, "PH.01.011": {"saldo": 6, "harga": 25000}, "PH.02.001": {"saldo": 10, "harga": 100000}, "PH.02.012": {"saldo": 7, "harga": 40000}, "PH.04.002": {"saldo": 4, "harga": 35000}, "PH.04.005": {"saldo": 5, "harga": 60000}, "PH.05.001": {"saldo": 5, "harga": 70000}, "PH.05.003": {"saldo": 9, "harga": 30000}, "PH.06.004": {"saldo": 5, "harga": 50000}, "PH.06.007": {"saldo": 18, "harga": 30000}, "PN.01.002": {"saldo": 3, "harga": 15000}, "PN.01.003": {"saldo": 2, "harga": 20000}, "PN.02.003": {"saldo": 17, "harga": 20000}, "PN.02.004": {"saldo": 9, "harga": 20000}, "BB.01.001": {"saldo": 9, "harga": 16000}};

const DEFAULT_KELOMPOK = {
  KH: { nama: "Karbohidrat", color: "#C48A00", bg: "#FBF1DA" },
  PH: { nama: "Protein Hewani", color: "#A8402A", bg: "#F8E6E1" },
  PN: { nama: "Protein Nabati", color: "#6E7A3B", bg: "#EEF0DD" },
  SY: { nama: "Sayuran", color: "#2B6E52", bg: "#DEEEE6" },
  BU: { nama: "Buah-buahan", color: "#C1602E", bg: "#F8E4D6" },
  BB: { nama: "Bahan Baku Lain", color: "#5B6472", bg: "#E7E9EC" },
};
// KELOMPOK sengaja bukan const: bisa ditambah kelompok baru saat runtime lewat menu Master Barang.
// eslint-disable-next-line import/no-mutable-exports
let KELOMPOK = { ...DEFAULT_KELOMPOK };
const KELOMPOK_PALETTE = [
  { color: "#C48A00", bg: "#FBF1DA" }, { color: "#A8402A", bg: "#F8E6E1" }, { color: "#6E7A3B", bg: "#EEF0DD" },
  { color: "#2B6E52", bg: "#DEEEE6" }, { color: "#C1602E", bg: "#F8E4D6" }, { color: "#5B6472", bg: "#E7E9EC" },
  { color: "#2F5D50", bg: "#DFEAE5" }, { color: "#8B4E9C", bg: "#EEE1F2" }, { color: "#2A6F8E", bg: "#DDEAF0" },
];
function nextKelompokPalette() {
  const used = Object.keys(KELOMPOK).length;
  return KELOMPOK_PALETTE[used % KELOMPOK_PALETTE.length];
}

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

/* ============================== STORAGE HELPERS ============================== */
const KEYS = {
  users: "sppg-users-v1",
  masuk: "sppg-transaksi-masuk-v1",
  keluar: "sppg-transaksi-keluar-v1",
  saldo: "sppg-saldo-per-periode-v1",
  master: "sppg-master-barang-v1",
  log: "sppg-log-aktivitas-v1",
  periode: "sppg-periode-v1",
  supplier: "sppg-supplier-v1",
  kelompok: "sppg-kelompok-v1",
};

const DEFAULT_SUPPLIERS = [];

const DEFAULT_USERS = {
  superadmin: { username: "superadmin", password: "super123" },
  admins: [{ id: "admin-demo", username: "admin1", password: "admin123", nama: "Admin Gudang 1" }],
  viewerPin: "1234",
};

const DEFAULT_PERIODE = {
  list: [{ id: "periode-1", nama: "Periode 1 - Februari 2026", mulai: "2026-02-16", selesai: "2026-02-28", closed: false }],
  activeId: "periode-1",
};
const DEFAULT_SALDO_BY_PERIODE = { "periode-1": SALDO_AWAL };

function fmtTgl(iso) {
  if (!iso) return "-";
  const d = new Date(iso + "T00:00:00");
  return d.toLocaleDateString("id-ID", { day: "2-digit", month: "short", year: "numeric" });
}

const WORKSPACE_ID = import.meta.env.VITE_WORKSPACE_ID || "gudang-utama";

async function storageGet(key, fallback) {
  try {
    const { data, error } = await supabase
      .from("app_storage_shared")
      .select("value")
      .eq("workspace_id", WORKSPACE_ID)
      .eq("key", key)
      .maybeSingle();
    if (error) throw error;
    return data?.value ?? fallback;
  } catch (e) {
    console.error("supabase storage get failed", key, e);
    return fallback;
  }
}
async function storageSet(key, value) {
  const { error } = await supabase
    .from("app_storage_shared")
    .upsert(
      { workspace_id: WORKSPACE_ID, key, value, updated_at: new Date().toISOString() },
      { onConflict: "workspace_id,key" },
    );
  if (error) {
    console.error("supabase storage set failed", key, error);
    throw error;
  }
}

/* ============================== ROOT APP ============================== */
export default function App() {
  const [loading, setLoading] = useState(true);
  const [users, setUsers] = useState(DEFAULT_USERS);
  const [masuk, setMasuk] = useState([]);
  const [keluar, setKeluar] = useState([]);
  const [saldoByPeriode, setSaldoByPeriode] = useState(DEFAULT_SALDO_BY_PERIODE);
  const [periodeState, setPeriodeState] = useState(DEFAULT_PERIODE);
  const [masterItems, setMasterItems] = useState(ITEMS);
  const [suppliers, setSuppliers] = useState(DEFAULT_SUPPLIERS);
  const [kelompokVersion, setKelompokVersion] = useState(0);
  const [log, setLog] = useState([]);
  const [session, setSession] = useState(null);
  const [page, setPage] = useState("dashboard");
  const [navOpen, setNavOpen] = useState(false);
  const [viewingPeriodeId, setViewingPeriodeId] = useState(null);

  useEffect(() => {
    (async () => {
      const [u, m, k, s, mi, lg, pr, sp, kl] = await Promise.all([
        storageGet(KEYS.users, null),
        storageGet(KEYS.masuk, null),
        storageGet(KEYS.keluar, null),
        storageGet(KEYS.saldo, null),
        storageGet(KEYS.master, null),
        storageGet(KEYS.log, null),
        storageGet(KEYS.periode, null),
        storageGet(KEYS.supplier, null),
        storageGet(KEYS.kelompok, null),
      ]);
      if (u) setUsers(u); else await storageSet(KEYS.users, DEFAULT_USERS);
      if (m) setMasuk(m);
      if (k) setKeluar(k);
      if (s) setSaldoByPeriode(s); else await storageSet(KEYS.saldo, DEFAULT_SALDO_BY_PERIODE);
      if (mi) setMasterItems(mi); else await storageSet(KEYS.master, ITEMS);
      if (lg) setLog(lg);
      if (sp) setSuppliers(sp); else await storageSet(KEYS.supplier, DEFAULT_SUPPLIERS);
      if (kl) { KELOMPOK = kl; setKelompokVersion((v) => v + 1); } else await storageSet(KEYS.kelompok, KELOMPOK);
      const finalPeriode = pr || DEFAULT_PERIODE;
      if (!pr) await storageSet(KEYS.periode, DEFAULT_PERIODE);
      setPeriodeState(finalPeriode);
      setViewingPeriodeId(finalPeriode.activeId);
      setLoading(false);
    })();
  }, []);

  const persist = {
    users: async (v) => { setUsers(v); await storageSet(KEYS.users, v); },
    masuk: async (v) => { setMasuk(v); await storageSet(KEYS.masuk, v); },
    keluar: async (v) => { setKeluar(v); await storageSet(KEYS.keluar, v); },
    saldoByPeriode: async (v) => { setSaldoByPeriode(v); await storageSet(KEYS.saldo, v); },
    master: async (v) => { setMasterItems(v); await storageSet(KEYS.master, v); },
    supplier: async (v) => { setSuppliers(v); await storageSet(KEYS.supplier, v); },
    log: async (v) => { setLog(v); await storageSet(KEYS.log, v); },
    periode: async (v) => { setPeriodeState(v); await storageSet(KEYS.periode, v); },
    kelompok: async (v) => { KELOMPOK = v; setKelompokVersion((x) => x + 1); await storageSet(KEYS.kelompok, v); },
  };

  const addLog = useCallback(async (aksi, detail) => {
    const entry = { id: uid(), waktu: new Date().toISOString(), oleh: session ? session.nama : "-", role: session ? session.role : "-", aksi, detail };
    const next = [entry, ...log].slice(0, 500);
    await persist.log(next);
  }, [log, session]);

  // pastikan nama supplier tersimpan di data supplier; jika belum ada, tambahkan otomatis
  const ensureSupplier = useCallback(async (namaMentah) => {
    const nama = (namaMentah || "").trim();
    if (!nama) return "-";
    const existing = suppliers.find((s) => s.nama.toLowerCase() === nama.toLowerCase());
    if (existing) return existing.nama;
    const baru = { id: uid(), nama };
    await persist.supplier([...suppliers, baru]);
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

  const tutupPeriode = useCallback(async ({ nama, mulai, selesai }) => {
    const activeId = periodeState.activeId;
    const hasilAkhir = computeStockFor(activeId);
    const newId = uid();
    const saldoBaru = {};
    for (const it of hasilAkhir) {
      saldoBaru[it.kode] = { saldo: it.saldoAkhir, harga: it.hargaTerakhir || 0 };
    }
    const nextPeriodeList = periodeState.list.map((p) => (p.id === activeId ? { ...p, closed: true } : p));
    nextPeriodeList.push({ id: newId, nama, mulai, selesai, closed: false });
    await persist.periode({ list: nextPeriodeList, activeId: newId });
    await persist.saldoByPeriode({ ...saldoByPeriode, [newId]: saldoBaru });
    setViewingPeriodeId(newId);
    await addLog("Tutup periode & buka periode baru", `${periodeState.list.find((p) => p.id === activeId)?.nama} -> ${nama}`);
  }, [periodeState, saldoByPeriode, computeStockFor, addLog]);

  const editPeriode = useCallback(async (id, patch) => {
    const nextList = periodeState.list.map((p) => (p.id === id ? { ...p, ...patch } : p));
    await persist.periode({ ...periodeState, list: nextList });
    await addLog("Ubah data periode", `${id}`);
  }, [periodeState, addLog]);

  if (loading || !viewingPeriodeId) {
    return (
      <div style={{ minHeight: "100vh", background: "#F1EEE4", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "Inter, sans-serif", color: "#6B6355" }}>
        Memuat data gudang...
      </div>
    );
  }

  if (!session) {
    return <LoginScreen users={users} onLogin={setSession} />;
  }

  const periode = {
    list: periodeState.list,
    activeId: periodeState.activeId,
    viewingId: viewingPeriodeId,
    setViewingId: setViewingPeriodeId,
    tutupPeriode,
    editPeriode,
  };

  const ctx = { session, stock, masuk, keluar, masterItems, saldoByPeriode, users, log, persist, addLog, periode, computeStockFor, suppliers, ensureSupplier, kelompokVersion };

  return (
    <div style={{ minHeight: "100vh", background: "#F1EEE4", fontFamily: "Inter, system-ui, sans-serif", color: "#20241F" }}>
      <GlobalStyle />
      <TopBar session={session} onLogout={() => setSession(null)} onMenu={() => setNavOpen((v) => !v)} />
      <PeriodeBar periode={periode} role={session.role} setPage={setPage} />
      {navOpen && <div className="nav-overlay show" onClick={() => setNavOpen(false)} />}
      <div style={{ display: "flex", maxWidth: 1280, margin: "0 auto" }}>
        <SideNav role={session.role} page={page} setPage={setPage} navOpen={navOpen} setNavOpen={setNavOpen} />
        <main style={{ flex: 1, padding: "20px 16px 60px", minWidth: 0 }}>
          <PageRouter page={page} ctx={ctx} />
        </main>
      </div>
      <BottomTabBar role={session.role} page={page} setPage={setPage} setNavOpen={setNavOpen} />
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

      /* ---- Native app feel on mobile: bottom tab bar ---- */
      .bottom-tabbar { display: none; }
      @media (max-width: 820px) {
        html { -webkit-tap-highlight-color: transparent; }
        button, .cekstok-pill, .side-nav button { -webkit-tap-highlight-color: transparent; }
        button:active { opacity: .7; }
        .bottom-tabbar {
          display: flex; position: fixed; left: 0; right: 0; bottom: 0; z-index: 50;
          background: #FBF9F3; border-top: 1px solid #E4DFCF;
          padding: 6px 4px calc(6px + env(safe-area-inset-bottom, 0px));
          box-shadow: 0 -4px 16px rgba(0,0,0,.06);
        }
        .bottom-tabbar-item {
          flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center;
          gap: 2px; background: none; border: none; padding: 6px 2px; cursor: pointer;
          color: #8B8371; font-size: 10px; font-weight: 600;
        }
        .bottom-tabbar-item.active { color: #2F5D50; }
        main { padding-bottom: 82px !important; }
      }
    `}</style>
  );
}

/* ============================== BOTTOM TAB BAR (mobile, native-app style) ============================== */
const TAB_KEYS = ["dashboard", "cekstok", "masuk", "keluar", "laporan"];
function BottomTabBar({ role, page, setPage, setNavOpen }) {
  const items = NAV_ITEMS.filter((i) => TAB_KEYS.includes(i.key) && i.roles.includes(role));
  return (
    <nav className="bottom-tabbar">
      {items.map((it) => {
        const Icon = it.icon;
        const active = page === it.key;
        return (
          <button key={it.key} className={"bottom-tabbar-item" + (active ? " active" : "")} onClick={() => setPage(it.key)}>
            <Icon size={19} strokeWidth={active ? 2.4 : 2} />
            {it.label.split(" ")[0]}
          </button>
        );
      })}
      <button className="bottom-tabbar-item" onClick={() => setNavOpen(true)}>
        <Menu size={19} />
        Menu
      </button>
    </nav>
  );
}

/* ============================== LOGiIN ============================== */
function LoginScreen({ users, onLogin }) {
  const [mode, setMode] = useState("staff");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [pin, setPin] = useState("");
  const [error, setError] = useState("");

  function submitStaff(e) {
    e.preventDefault();
    setError("");
    if (username === users.superadmin.username && password === users.superadmin.password) {
      onLogin({ role: "superadmin", nama: "Super Admin", username });
      return;
    }
    const admin = users.admins.find((a) => a.username === username && a.password === password);
    if (admin) {
      onLogin({ role: "admin", nama: admin.nama, username: admin.username, id: admin.id });
      return;
    }
    setError("Username atau password salah.");
  }

  function submitPin(e) {
    e.preventDefault();
    setError("");
    if (pin === users.viewerPin) {
      onLogin({ role: "viewer", nama: "Viewer", username: "viewer" });
      return;
    }
    setError("PIN salah.");
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
            <button onClick={() => { setMode("staff"); setError(""); }} style={{ flex: 1, padding: "9px 0", borderRadius: 7, border: "none", cursor: "pointer", fontWeight: 600, fontSize: 13.5, background: mode === "staff" ? "#2F5D50" : "transparent", color: mode === "staff" ? "#fff" : "#5C5646" }}>Admin</button>
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
              <button type="submit" style={primaryBtn}><LogIn size={16} /> Masuk</button>
            </form>
          ) : (
            <form onSubmit={submitPin}>
              <Field label="Masukkan PIN">
                <input value={pin} onChange={(e) => setPin(e.target.value.replace(/\D/g, ""))} inputMode="numeric" maxLength={8} autoFocus style={{ ...inputStyle, textAlign: "center", letterSpacing: 6, fontSize: 20 }} placeholder="----" />
              </Field>
              {error && <ErrorNote text={error} />}
              <button type="submit" style={primaryBtn}><Eye size={16} /> Lihat Laporan</button>
            </form>
          )}
          <div style={{ marginTop: 16, fontSize: 11.5, color: "#8B8371", lineHeight: 1.5 }}>
          
          </div>
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
  { key: "cekstok", label: "Cek Stok", icon: Eye, roles: ["superadmin", "admin", "viewer"] },
  { key: "masuk", label: "Penerimaan Barang", icon: ArrowDownToLine, roles: ["superadmin", "admin"] },
  { key: "keluar", label: "Pengeluaran Barang", icon: ArrowUpFromLine, roles: ["superadmin", "admin"] },
  { key: "laporan", label: "Laporan Stock", icon: ClipboardList, roles: ["superadmin", "admin", "viewer"] },
  { key: "master", label: "Master Barang", icon: Package, roles: ["superadmin"] },
  { key: "supplier", label: "Data Supplier", icon: Truck, roles: ["superadmin", "admin"] },
  { key: "periode", label: "Pengaturan Periode", icon: CalendarRange, roles: ["superadmin"] },
  { key: "users", label: "Manajemen User", icon: Users, roles: ["superadmin"] },
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
    case "cekstok": return <CekStok ctx={ctx} />;
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

/* ============================== CEK STOK (realtime, tanpa rincian masuk/keluar) ============================== */
function CekStok({ ctx }) {
  const { stock, periode } = ctx;
  const [q, setQ] = useState("");
  const [kel, setKel] = useState("ALL");
  const periodeInfo = periode.list.find((p) => p.id === periode.viewingId);

  const rows = useMemo(() => stock.filter((it) => {
    if (kel !== "ALL" && kelompokOf(it.kode) !== kel) return false;
    if (q && !(it.nama.toLowerCase().includes(q.toLowerCase()) || it.kode.toLowerCase().includes(q.toLowerCase()))) return false;
    return true;
  }).sort((a, b) => a.kode.localeCompare(b.kode)), [stock, kel, q]);

  function statusOf(it) {
    if (it.saldoAkhir <= 0) return { label: "Habis", color: "#A8402A", bg: "#F8E4E0" };
    if (it.saldoAkhir <= 3) return { label: "Menipis", color: "#C1602E", bg: "#FBEADD" };
    return { label: "Aman", color: "#2F5D50", bg: "#E4EFE9" };
  }

  return (
    <div>
      <PageHeader title="Cek Stok" subtitle={`Stok realtime terkini - ${periodeInfo ? periodeInfo.nama : ""}. Tanpa rincian masuk/keluar, langsung lihat sisa stok saat ini.`} />
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
            <thead><tr>{["Kode", "Nama Barang", "Satuan", "Stok Saat Ini", "Status"].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
            <tbody>
              {rows.map((it) => {
                const st = statusOf(it);
                return (
                  <tr key={it.kode}>
                    <Td className="mono">{it.kode}</Td>
                    <Td>{it.nama}</Td>
                    <Td>{it.satuan}</Td>
                    <Td className="mono" style={{ fontWeight: 700, fontSize: 14.5, color: st.color }}>{it.saldoAkhir}</Td>
                    <Td><span style={{ fontSize: 10.5, fontWeight: 700, color: st.color, background: st.bg, padding: "3px 8px", borderRadius: 5 }}>{st.label}</span></Td>
                  </tr>
                );
              })}
              {rows.length === 0 && <tr><td colSpan={5}><Empty text="Tidak ada barang yang cocok." /></td></tr>}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
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
    const entry = {
      id: uid(), tanggal: form.tanggal, supplier: namaSupplier, kode: item.kode, nama: item.nama, satuan: item.satuan,
      vol: Number(form.vol), harga: Number(form.harga) || 0, jumlah: (Number(form.vol) || 0) * (Number(form.harga) || 0),
      oleh: session.nama, waktu: new Date().toISOString(), periodeId: periode.activeId,
    };
    const next = [entry, ...masuk];
    await persist.masuk(next);
    await addLog("Penerimaan barang", `${item.nama} +${entry.vol} ${item.satuan}`);
    setForm({ tanggal: todayInput(), supplier: "", kode: "", vol: "", harga: "" });
    setMsg("Tersimpan.");
    setTimeout(() => setMsg(""), 2000);
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
    const entry = { id: uid(), tanggal: form.tanggal, kode: item.kode, nama: item.nama, satuan: item.satuan, vol: Number(form.vol), oleh: session.nama, waktu: new Date().toISOString(), periodeId: periode.activeId };
    await persist.keluar([entry, ...keluar]);
    await addLog("Pengeluaran barang", `${item.nama} -${entry.vol} ${item.satuan}`);
    setForm({ tanggal: todayInput(), kode: "", vol: "" });
    setMsg("Tersimpan.");
    setTimeout(() => setMsg(""), 2000);
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
  const { masterItems, persist, addLog, saldoByPeriode, periode, kelompokVersion } = ctx;
  const saldoAwal = saldoByPeriode[periode.activeId] || {};
  const periodeAktifInfo = periode.list.find((p) => p.id === periode.activeId);
  const [q, setQ] = useState("");
  const kelompokOptions = Object.keys(KELOMPOK);
  const [form, setForm] = useState({ kelompokKode: kelompokOptions[0] || "", kodeSub: "", nama: "", satuan: "", saldo: "", harga: "" });
  const [editKode, setEditKode] = useState(null);
  const [msg, setMsg] = useState("");
  const [showNewKelompok, setShowNewKelompok] = useState(false);
  const [newKelompok, setNewKelompok] = useState({ kode: "", nama: "" });
  const [kelMsg, setKelMsg] = useState("");

  const filtered = masterItems.filter((i) => !q || i.nama.toLowerCase().includes(q.toLowerCase()) || i.kode.toLowerCase().includes(q.toLowerCase())).sort((a, b) => a.kode.localeCompare(b.kode));

  function buildKode() {
    return `${form.kelompokKode}.${form.kodeSub}`.replace(/\.+$/, "");
  }

  async function addItem(e) {
    e.preventDefault();
    const kode = buildKode();
    if (!form.kelompokKode || !form.kodeSub || !form.nama || !form.satuan) { setMsg("Kelompok, kode, nama, dan satuan wajib diisi."); return; }
    if (masterItems.some((i) => i.kode === kode) && !editKode) { setMsg("Kode sudah dipakai."); return; }
    let next;
    if (editKode) {
      next = masterItems.map((i) => (i.kode === editKode ? { kode, nama: form.nama, satuan: form.satuan } : i));
    } else {
      next = [...masterItems, { kode, nama: form.nama, satuan: form.satuan }];
    }
    await persist.master(next);
    if (form.saldo || form.harga) {
      await persist.saldoByPeriode({ ...saldoByPeriode, [periode.activeId]: { ...saldoAwal, [kode]: { saldo: Number(form.saldo) || 0, harga: Number(form.harga) || 0 } } });
    }
    await addLog(editKode ? "Ubah master barang" : "Tambah master barang", `${kode} - ${form.nama}`);
    setForm({ kelompokKode: form.kelompokKode, kodeSub: "", nama: "", satuan: "", saldo: "", harga: "" });
    setEditKode(null);
    setMsg("Tersimpan.");
    setTimeout(() => setMsg(""), 2000);
  }

  function editRow(it) {
    const sa = saldoAwal[it.kode] || {};
    const kelompokKode = kelompokOf(it.kode);
    const kodeSub = it.kode.slice(kelompokKode.length + 1);
    setEditKode(it.kode);
    setForm({ kelompokKode, kodeSub, nama: it.nama, satuan: it.satuan, saldo: sa.saldo || "", harga: sa.harga || "" });
  }

  async function removeItem(kode) {
    if (!confirm("Hapus barang ini dari master data?")) return;
    await persist.master(masterItems.filter((i) => i.kode !== kode));
    await addLog("Hapus master barang", kode);
  }

  async function addKelompokBaru(e) {
    e.preventDefault();
    const kode = newKelompok.kode.trim().toUpperCase();
    if (!kode || kode.length < 2 || kode.length > 4) { setKelMsg("Kode kelompok 2-4 huruf, mis. LN."); return; }
    if (!newKelompok.nama.trim()) { setKelMsg("Nama kelompok wajib diisi."); return; }
    if (KELOMPOK[kode]) { setKelMsg("Kode kelompok sudah dipakai."); return; }
    const palette = nextKelompokPalette();
    await persist.kelompok({ ...KELOMPOK, [kode]: { nama: newKelompok.nama.trim(), color: palette.color, bg: palette.bg } });
    await addLog("Tambah kelompok baru", `${kode} - ${newKelompok.nama.trim()}`);
    setForm((f) => ({ ...f, kelompokKode: kode }));
    setNewKelompok({ kode: "", nama: "" });
    setShowNewKelompok(false);
    setKelMsg("");
  }

  return (
    <div>
      <PageHeader title="Master Barang" subtitle={`Kelola daftar barang persediaan. Saldo awal berlaku untuk periode aktif: ${periodeAktifInfo ? periodeAktifInfo.nama : "-"}.`} />
      <div className="grid-master">
        <Card>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10 }}>{editKode ? "Ubah Barang" : "Tambah Barang"}</div>
          <form onSubmit={addItem}>
            <Field label="Kelompok Barang">
              <div style={{ display: "flex", gap: 8 }}>
                <select value={form.kelompokKode} onChange={(e) => setForm({ ...form, kelompokKode: e.target.value })} style={{ ...inputStyle, flex: 1 }}>
                  {kelompokOptions.map((k) => <option key={k} value={k}>{k} - {KELOMPOK[k].nama}</option>)}
                </select>
                <button type="button" onClick={() => setShowNewKelompok((v) => !v)} title="Tambah kelompok baru" style={{ ...iconBtn, padding: "0 12px" }}>
                  <Plus size={15} />
                </button>
              </div>
            </Field>
            {showNewKelompok && (
              <div style={{ background: "#F1EEE4", border: "1px dashed #DAD4C2", borderRadius: 8, padding: 12, marginBottom: 14 }}>
                <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 8, color: "#5C5646" }}>Tambah Kelompok Baru</div>
                <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                  <input value={newKelompok.kode} onChange={(e) => setNewKelompok({ ...newKelompok, kode: e.target.value.toUpperCase() })} placeholder="Kode (mis. LN)" maxLength={4} style={{ ...inputStyle, width: 100 }} />
                  <input value={newKelompok.nama} onChange={(e) => setNewKelompok({ ...newKelompok, nama: e.target.value })} placeholder="Nama kelompok" style={{ ...inputStyle, flex: 1 }} />
                </div>
                {kelMsg && <div style={{ fontSize: 12, color: "#A8402A", marginBottom: 8 }}>{kelMsg}</div>}
                <button type="button" onClick={addKelompokBaru} style={{ ...primaryBtn, width: "auto", padding: "8px 14px", marginTop: 0 }}>Simpan Kelompok</button>
              </div>
            )}
            <Field label="Kode Barang">
              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <span className="mono" style={{ fontSize: 13, fontWeight: 700, color: "#8B8371", flexShrink: 0 }}>{form.kelompokKode}.</span>
                <input value={form.kodeSub} onChange={(e) => setForm({ ...form, kodeSub: e.target.value })} style={inputStyle} placeholder="mis. 01.010" disabled={!!editKode} />
              </div>
            </Field>
            <Field label="Nama Barang"><input value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} style={inputStyle} /></Field>
            <Field label="Satuan"><input value={form.satuan} onChange={(e) => setForm({ ...form, satuan: e.target.value })} style={inputStyle} placeholder="kg / pcs / liter" /></Field>
            <div style={{ display: "flex", gap: 10 }}>
              <div style={{ flex: 1 }}><Field label={`Saldo Awal (${periodeAktifInfo ? periodeAktifInfo.nama : "periode aktif"})`}><input type="number" value={form.saldo} onChange={(e) => setForm({ ...form, saldo: e.target.value })} style={inputStyle} /></Field></div>
              <div style={{ flex: 1 }}><Field label="Harga Awal"><input type="number" value={form.harga} onChange={(e) => setForm({ ...form, harga: e.target.value })} style={inputStyle} /></Field></div>
            </div>
            {msg && <div style={{ fontSize: 12.5, marginBottom: 10, color: msg.includes("wajib") || msg.includes("dipakai") ? "#A8402A" : "#2F5D50" }}>{msg}</div>}
            <div style={{ display: "flex", gap: 8 }}>
              <button type="submit" style={primaryBtn}>{editKode ? <><Pencil size={15} /> Simpan Perubahan</> : <><Plus size={15} /> Tambah Barang</>}</button>
              {editKode && <button type="button" onClick={() => { setEditKode(null); setForm({ kelompokKode: kelompokOptions[0] || "", kodeSub: "", nama: "", satuan: "", saldo: "", harga: "" }); }} style={{ ...primaryBtn, background: "#EDE8DA", color: "#3A4038" }}><X size={15} /></button>}
            </div>
          </form>
        </Card>
        <Card style={{ maxHeight: 560, overflowY: "auto" }}>
          <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Cari barang..." style={{ ...inputStyle, marginBottom: 10 }} />
          <div className="table-wrap">
            <table>
              <thead><tr>{["Kelompok", "Kode", "Nama", "Satuan", "Saldo Awal", ""].map((h) => <Th key={h}>{h}</Th>)}</tr></thead>
              <tbody>
                {filtered.map((it) => (
                  <tr key={it.kode}>
                    <Td><KelompokChip kode={it.kode} /></Td>
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
    let next;
    if (editId) {
      next = suppliers.map((s) => (s.id === editId ? { ...s, nama, kontak: form.kontak } : s));
    } else {
      next = [...suppliers, { id: uid(), nama, kontak: form.kontak }];
    }
    await persist.supplier(next);
    await addLog(editId ? "Ubah data supplier" : "Tambah supplier", nama);
    setForm({ nama: "", kontak: "" });
    setEditId(null);
    setMsg("Tersimpan.");
    setTimeout(() => setMsg(""), 2000);
  }

  function editRow(s) {
    setEditId(s.id);
    setForm({ nama: s.nama, kontak: s.kontak || "" });
  }

  async function removeRow(s) {
    if (!confirm(`Hapus supplier "${s.nama}"? Riwayat transaksi lama tidak akan terhapus.`)) return;
    await persist.supplier(suppliers.filter((x) => x.id !== s.id));
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
  const { users, persist, addLog } = ctx;
  const [form, setForm] = useState({ nama: "", username: "", password: "" });
  const [spForm, setSpForm] = useState({ username: users.superadmin.username, password: "" });
  const [pinForm, setPinForm] = useState(users.viewerPin);
  const [msg, setMsg] = useState("");

  async function addAdmin(e) {
    e.preventDefault();
    if (!form.nama || !form.username || !form.password) { setMsg("Lengkapi semua kolom."); return; }
    if (users.admins.some((a) => a.username === form.username)) { setMsg("Username sudah dipakai."); return; }
    const next = { ...users, admins: [...users.admins, { id: uid(), ...form }] };
    await persist.users(next);
    await addLog("Tambah admin", `${form.nama} (${form.username})`);
    setForm({ nama: "", username: "", password: "" });
    setMsg("Admin ditambahkan.");
    setTimeout(() => setMsg(""), 2000);
  }

  async function removeAdmin(id, nama) {
    if (!confirm(`Hapus akun admin "${nama}"?`)) return;
    await persist.users({ ...users, admins: users.admins.filter((a) => a.id !== id) });
    await addLog("Hapus admin", nama);
  }

  async function saveSuperAdmin(e) {
    e.preventDefault();
    const next = { ...users, superadmin: { username: spForm.username, password: spForm.password || users.superadmin.password } };
    await persist.users(next);
    await addLog("Ubah kredensial super admin", spForm.username);
    setSpForm({ ...spForm, password: "" });
    setMsg("Kredensial super admin diperbarui.");
    setTimeout(() => setMsg(""), 2000);
  }

  async function savePin(e) {
    e.preventDefault();
    await persist.users({ ...users, viewerPin: pinForm });
    await addLog("Ubah PIN viewer", "***");
    setMsg("PIN viewer diperbarui.");
    setTimeout(() => setMsg(""), 2000);
  }

  return (
    <div>
      <PageHeader title="Manajemen User" subtitle="Kelola akun Admin, kredensial Super Admin, dan PIN Viewer." />
      {msg && <div style={{ fontSize: 12.5, color: "#2F5D50", marginBottom: 14 }}>{msg}</div>}
      <div className="grid-eq2">
        <Card>
          <div className="sg" style={{ fontWeight: 700, marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}><Users size={15} /> Akun Admin</div>
          <form onSubmit={addAdmin} style={{ marginBottom: 14, display: "grid", gap: 8 }}>
            <input value={form.nama} onChange={(e) => setForm({ ...form, nama: e.target.value })} placeholder="Nama lengkap" style={inputStyle} />
            <input value={form.username} onChange={(e) => setForm({ ...form, username: e.target.value })} placeholder="Username" style={inputStyle} />
            <input value={form.password} onChange={(e) => setForm({ ...form, password: e.target.value })} placeholder="Password" style={inputStyle} />
            <button type="submit" style={{ ...primaryBtn, width: "auto", padding: "9px 16px" }}><Plus size={15} /> Tambah Admin</button>
          </form>
          {users.admins.map((a) => (
            <div key={a.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "8px 0", borderBottom: "1px solid #EEE9DB" }}>
              <div>
                <div style={{ fontSize: 13, fontWeight: 600 }}>{a.nama}</div>
                <div style={{ fontSize: 11.5, color: "#8B8371" }} className="mono">{a.username}</div>
              </div>
              <button onClick={() => removeAdmin(a.id, a.nama)} style={{ ...iconBtn, color: "#A8402A" }}><Trash2 size={13} /></button>
            </div>
          ))}
          {users.admins.length === 0 && <Empty text="Belum ada akun admin." />}
        </Card>

        <div style={{ display: "grid", gap: 16 }}>
          <Card>
            <div className="sg" style={{ fontWeight: 700, marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}><ShieldCheck size={15} /> Kredensial Super Admin</div>
            <form onSubmit={saveSuperAdmin} style={{ display: "grid", gap: 8 }}>
              <Field label="Username"><input value={spForm.username} onChange={(e) => setSpForm({ ...spForm, username: e.target.value })} style={inputStyle} /></Field>
              <Field label="Password baru (kosongkan jika tidak diubah)"><input type="password" value={spForm.password} onChange={(e) => setSpForm({ ...spForm, password: e.target.value })} style={inputStyle} /></Field>
              <button type="submit" style={{ ...primaryBtn, width: "auto", padding: "9px 16px" }}>Simpan</button>
            </form>
          </Card>
          <Card>
            <div className="sg" style={{ fontWeight: 700, marginBottom: 10, display: "flex", alignItems: "center", gap: 6 }}><KeyRound size={15} /> PIN Viewer</div>
            <form onSubmit={savePin} style={{ display: "grid", gap: 8 }}>
              <Field label="PIN"><input value={pinForm} onChange={(e) => setPinForm(e.target.value.replace(/\D/g, ""))} inputMode="numeric" style={inputStyle} /></Field>
              <button type="submit" style={{ ...primaryBtn, width: "auto", padding: "9px 16px" }}>Simpan PIN</button>
            </form>
          </Card>
          <div style={{ fontSize: 11.5, color: "#8B8371", lineHeight: 1.5, padding: "0 4px" }}>
            Catatan keamanan: aplikasi ini menyimpan data operasional di Supabase. Akun pemilik dilindungi Supabase Auth, sedangkan akun internal digunakan untuk pembagian akses di dalam aplikasi.
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
