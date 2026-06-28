import { useState, useMemo, useEffect } from "react";
import { Search, ShoppingCart, X, Plus, Minus, LogOut, Leaf, ChevronLeft, Trash2, Check, ZoomIn, Send,
  ShieldCheck, Download, Upload, Package, ClipboardList, Sprout, AlertTriangle, CheckCircle2 } from "lucide-react";
import * as XLSX from "xlsx";
import { auth, db } from "./firebase";
import { onAuthStateChanged, createUserWithEmailAndPassword, signInWithEmailAndPassword, signOut } from "firebase/auth";
import { doc, getDoc, setDoc, updateDoc, collection, addDoc, onSnapshot, query, orderBy, serverTimestamp, runTransaction } from "firebase/firestore";

const PRODUCTS = [{"id": 0, "code": "CCBJ", "name": "Capsicum Bhut Jolokia Red", "cat": "Capsicums", "img": "https://i.ibb.co/k2F22zB4/Capsicum-Bhut-Jolokia-Red.png", "stock": 50}, {"id": 1, "code": "CCJa", "name": "Capsicum Jalapeno", "cat": "Capsicums", "img": "https://i.ibb.co/JFSsrJdM/Capsicum-Jalapeno.png", "stock": 50}, {"id": 2, "code": "CCBE", "name": "Capsicum Bird's Eye", "cat": "Capsicums", "img": "https://i.ibb.co/v4Jc3FX3/Capsicum-Birds-Eye.png", "stock": 100}, {"id": 3, "code": "CCPa", "name": "Capsicum Paprika", "cat": "Capsicums", "img": "https://i.ibb.co/7tTwzg8F/Capsicum-Paprika.png", "stock": 50}, {"id": 4, "code": "CCBC", "name": "Capsicum Bishop's Crown", "cat": "Capsicums", "img": "https://i.ibb.co/1Y2Tfc7s/Capsicum-Bishops-Crown.png", "stock": 50}, {"id": 5, "code": "CCPe", "name": "Capsicum Peppadew", "cat": "Capsicums", "img": "https://i.ibb.co/HfXtxQnw/Capsicum-Peppadew.png", "stock": 100}, {"id": 6, "code": "CCWo", "name": "Capsicum California Wonder", "cat": "Capsicums", "img": "https://i.ibb.co/0jPmWqw8/Capsicum-California-Wonder.png", "stock": 100}, {"id": 7, "code": "CCPo", "name": "Capsicum Polombo", "cat": "Capsicums", "img": "https://i.ibb.co/GvJ9W5cf/Capsicum-Polombo.png", "stock": 50}, {"id": 8, "code": "CCCR", "name": "Capsicum Carolina Reaper", "cat": "Capsicums", "img": "https://i.ibb.co/NnfyWsRN/Capsicum-Carolina-Reaper.png", "stock": 50}, {"id": 9, "code": "CCST", "name": "Capsicum Scorpion Tongue Black", "cat": "Capsicums", "img": "https://i.ibb.co/3Y9LPxHc/Capsicum-Scorpion-Tongue-Black.png", "stock": 50}, {"id": 10, "code": "CCCa", "name": "Capsicum Cayenne", "cat": "Capsicums", "img": "https://i.ibb.co/Kczfvgwp/Capsicum-Cayenne.png", "stock": 50}, {"id": 11, "code": "CCSB", "name": "Capsicum Scotch Bonnet", "cat": "Capsicums", "img": "https://i.ibb.co/MxS84vYM/Capsicum-Scotch-Bonnet.png", "stock": 50}, {"id": 12, "code": "CCCH", "name": "Capsicum Chocolate Habanero", "cat": "Capsicums", "img": "https://i.ibb.co/k2kFY3TL/Capsicum-Chocolate-Habanero.png", "stock": 50}, {"id": 13, "code": "CCTM", "name": "Capsicum Trinidad Moruga Scorpion", "cat": "Capsicums", "img": "https://i.ibb.co/WW7dRQr7/Capsicum-Trinidad-Scropion-Moruga.png", "stock": 50}, {"id": 14, "code": "CCHO", "name": "Capsicum Habanero Orange", "cat": "Capsicums", "img": "https://i.ibb.co/CsS3g3Sm/Capsicum-Habanero-Orange.png", "stock": 100}, {"id": 15, "code": "LDWL", "name": "Dichondra Wonder Lawn", "cat": "Crop Cover & Lawns", "img": "https://i.ibb.co/zHsTVWGg/Dichondra-Wonder-Lawn.png", "stock": 50}, {"id": 16, "code": "CPPCd", "name": "Poultry Pasture Cool Season Mix 1kg", "cat": "Crop Cover & Lawns", "img": null, "stock": 10}, {"id": 17, "code": "CCIH", "name": "Industrial Hemp", "cat": "Crop Cover & Lawns", "img": "https://i.ibb.co/C5TS7dDh/Hemp-Industrial-Hemp.png", "stock": 50}, {"id": 18, "code": "CPPW", "name": "Poultry Pasture Warm Season Mix", "cat": "Crop Cover & Lawns", "img": "https://i.ibb.co/WrGtzbR/Pasture-Mix-Poultry-Pasture-Mix.jpg", "stock": 100}, {"id": 19, "code": "LLML", "name": "LM Lawn", "cat": "Crop Cover & Lawns", "img": "https://i.ibb.co/VYzw44GQ/Lawn-LM-Lawn.png", "stock": 0}, {"id": 20, "code": "CPPWa", "name": "Poultry Pasture Warm Season Mix 100g", "cat": "Crop Cover & Lawns", "img": null, "stock": 10}, {"id": 21, "code": "CCLE", "name": "Loofah European", "cat": "Crop Cover & Lawns", "img": "https://i.ibb.co/WNYNHGSH/Loofah-European.png", "stock": 20}, {"id": 22, "code": "CPPWb", "name": "Poultry Pasture Warm Season Mix 250g", "cat": "Crop Cover & Lawns", "img": null, "stock": 10}, {"id": 23, "code": "LPGr", "name": "Pets Grass", "cat": "Crop Cover & Lawns", "img": "https://i.ibb.co/9H3VPWt9/Grass-Pets-Grass.png", "stock": 50}, {"id": 24, "code": "CPPWc", "name": "Poultry Pasture Warm Season Mix 500g", "cat": "Crop Cover & Lawns", "img": null, "stock": 10}, {"id": 25, "code": "CPPC", "name": "Poultry Pasture Cool Season Mix", "cat": "Crop Cover & Lawns", "img": null, "stock": 100}, {"id": 26, "code": "CPPWd", "name": "Poultry Pasture Warm Season Mix 1kg", "cat": "Crop Cover & Lawns", "img": null, "stock": 10}, {"id": 27, "code": "CPPCa", "name": "Poultry Pasture Cool Season Mix 100g", "cat": "Crop Cover & Lawns", "img": null, "stock": 10}, {"id": 28, "code": "CCRs", "name": "Rapeseed", "cat": "Crop Cover & Lawns", "img": "https://i.ibb.co/bRYYtDpd/Brassica-Rape-Seed.png", "stock": 100}, {"id": 29, "code": "CPPCb", "name": "Poultry Pasture Cool Season Mix 250g", "cat": "Crop Cover & Lawns", "img": null, "stock": 10}, {"id": 30, "code": "CTVG", "name": "Tobacco Virginia Gold", "cat": "Crop Cover & Lawns", "img": "https://i.ibb.co/9kPRRr7S/Tobacco.jpg", "stock": 50}, {"id": 31, "code": "CPPCc", "name": "Poultry Pasture Cool Season Mix 500g", "cat": "Crop Cover & Lawns", "img": null, "stock": 10}, {"id": 32, "code": "FABl", "name": "Agapanthus Blue", "cat": "Flowers", "img": "https://i.ibb.co/dJgKb5xt/Agapanthus-Blue.png", "stock": 100}, {"id": 33, "code": "FLar", "name": "Larkspur", "cat": "Flowers", "img": "https://i.ibb.co/Kj1XLS12/Larkspur-Mix.png", "stock": 100}, {"id": 34, "code": "FAWh", "name": "Alyssum White", "cat": "Flowers", "img": "https://i.ibb.co/FLpSBLhT/Alyssum-White.png", "stock": 100}, {"id": 35, "code": "FLob", "name": "Lobelia", "cat": "Flowers", "img": "https://i.ibb.co/848RfBzS/Lobelia.png", "stock": 100}, {"id": 36, "code": "FAMa", "name": "Ammi Majus", "cat": "Flowers", "img": "https://i.ibb.co/C3w0bfNZ/Ammi-Majus.png", "stock": 100}, {"id": 37, "code": "FMAf", "name": "Marigold African", "cat": "Flowers", "img": "https://i.ibb.co/sdZMSpJL/Marigold-African.png", "stock": 100}, {"id": 38, "code": "Faqu", "name": "Aquilegia", "cat": "Flowers", "img": "https://i.ibb.co/WvYrQYTj/Aquilegia-Mix.png", "stock": 100}, {"id": 39, "code": "FMBR", "name": "Marigold Bronza Red", "cat": "Flowers", "img": "https://i.ibb.co/fzsw7w5Q/Marigold-Bronze-Red.png", "stock": 100}, {"id": 40, "code": "FBal", "name": "Balsam", "cat": "Flowers", "img": "https://i.ibb.co/qMSxGXXS/Balsam-Mix.png", "stock": 100}, {"id": 41, "code": "FMYe", "name": "Marigold Yellow", "cat": "Flowers", "img": "https://i.ibb.co/mrNvHnHJ/Marigold-Yellow.png", "stock": 100}, {"id": 42, "code": "FCMi", "name": "Calendula Mix", "cat": "Flowers", "img": "https://i.ibb.co/s9TXHfXn/Calendula-Mix.png", "stock": 100}, {"id": 43, "code": "FMBV", "name": "Mesembryanthemum Bokbaai Vygie", "cat": "Flowers", "img": "https://i.ibb.co/dwyYSjj2/Mesembryanthem-Bokbaai-Vygie-Mix.png", "stock": 100}, {"id": 44, "code": "FCel", "name": "Celosia", "cat": "Flowers", "img": "https://i.ibb.co/rGyZnwN1/Celosia-Mix.png", "stock": 100}, {"id": 45, "code": "FNAM", "name": "Nasturtium Alaska Mix", "cat": "Flowers", "img": "https://i.ibb.co/svvLgHBG/Nasturtium-Alaska-Mix.png", "stock": 100}, {"id": 46, "code": "FCSD", "name": "Cineraria Silverdust Dusty Miller", "cat": "Flowers", "img": "https://i.ibb.co/HTjBW5hG/Dusty-Miller-Cinematic-Silverdust.png", "stock": 100}, {"id": 47, "code": "FOst", "name": "Osteospermum", "cat": "Flowers", "img": "https://i.ibb.co/9344ypMX/Osteospermum-Mix.png", "stock": 100}, {"id": 48, "code": "FCle", "name": "Cleome", "cat": "Flowers", "img": "https://i.ibb.co/mFPk2Lzd/Cleome-Mix.png", "stock": 0}, {"id": 49, "code": "FPBD", "name": "Poppy Black Dragon", "cat": "Flowers", "img": "https://i.ibb.co/Xxq6k48Q/Poppy-Black-Dragon.png", "stock": 100}, {"id": 50, "code": "FCBB", "name": "Cornflower Blue Boy", "cat": "Flowers", "img": "https://i.ibb.co/jkLhqR4Z/Cornflower-Blue-Boy.png", "stock": 100}, {"id": 51, "code": "FPCa", "name": "Poppy California", "cat": "Flowers", "img": "https://i.ibb.co/67PF5ZXW/Poppy-California.png", "stock": 100}, {"id": 52, "code": "FCLa", "name": "Cosmos Laced", "cat": "Flowers", "img": "https://i.ibb.co/9mdB4J5k/Cosmos-Laced.png", "stock": 100}, {"id": 53, "code": "FPFR", "name": "Poppy Flanders Red", "cat": "Flowers", "img": "https://i.ibb.co/v4g1SxWb/Poppy-Flanders-Red.png", "stock": 100}, {"id": 54, "code": "FCSS", "name": "Cosmos Sea Shell Pink", "cat": "Flowers", "img": "https://i.ibb.co/r25zRzVK/Cosmos-Seashell-Pink.png", "stock": 100}, {"id": 55, "code": "FPMi", "name": "Poppy Mix", "cat": "Flowers", "img": "https://i.ibb.co/ZRwzB1BW/Poppy-Mix.png", "stock": 100}, {"id": 56, "code": "FCSM", "name": "Cosmos Sensation Mix", "cat": "Flowers", "img": "https://i.ibb.co/whymqWK6/Cosmos-Sensation-Mix.png", "stock": 100}, {"id": 57, "code": "FPPL", "name": "Poppy Peony Lavender", "cat": "Flowers", "img": "https://i.ibb.co/tMjbRDh4/Poppy-Peony-Lavender.png", "stock": 0}, {"id": 58, "code": "FCVe", "name": "Cosmos Veldfire", "cat": "Flowers", "img": "https://i.ibb.co/RpL0R9wL/Cosmos-Veldfire.png", "stock": 100}, {"id": 59, "code": "FPPM", "name": "Poppy Peony Mix", "cat": "Flowers", "img": "https://i.ibb.co/xtyD4sMM/Poppy-Peony-Mix.png", "stock": 100}, {"id": 60, "code": "FDah", "name": "Dahlia", "cat": "Flowers", "img": "https://i.ibb.co/tMqdkvvk/Dahlia-Mix.png", "stock": 100}, {"id": 61, "code": "FPPPi", "name": "Poppy Peony Pink", "cat": "Flowers", "img": "https://i.ibb.co/7FTD9cS/Poppy-Peony-Pink.png", "stock": 100}, {"id": 62, "code": "FDel", "name": "Delphinium", "cat": "Flowers", "img": "https://i.ibb.co/gbtpCfd1/Delphinium.png", "stock": 100}, {"id": 63, "code": "FPPPu", "name": "Poppy Peony Purple", "cat": "Flowers", "img": "https://i.ibb.co/pG2vwcd/Poppy-Peony-Purple.png", "stock": 100}, {"id": 64, "code": "FDMi", "name": "Dianthus", "cat": "Flowers", "img": "https://i.ibb.co/Pz5tXS98/Dianthus-Mix.png", "stock": 100}, {"id": 65, "code": "FPPR", "name": "Poppy Peony Red", "cat": "Flowers", "img": "https://i.ibb.co/nZch0DZ/Poppy-Peony-Red.png", "stock": 100}, {"id": 66, "code": "FDAN", "name": "Dimorphotheca African/Namakwal. Daisy", "cat": "Flowers", "img": "https://i.ibb.co/jcN8PrZ/Dimorphotheca-African-Namaqualand-Daisy.png", "stock": 100}, {"id": 67, "code": "FPPRe", "name": "Poppy Pepperbox Red", "cat": "Flowers", "img": "https://i.ibb.co/Q38FGJtQ/Poppy-Pepperbox-Red.png", "stock": 100}, {"id": 68, "code": "FEWh", "name": "Erigeron White", "cat": "Flowers", "img": "https://i.ibb.co/Gfh7wT8w/Erigeron-White.png", "stock": 100}, {"id": 69, "code": "FPPi", "name": "Poppy Pink", "cat": "Flowers", "img": "https://i.ibb.co/LXQbZXdn/Poppy-Pink.png", "stock": 100}, {"id": 70, "code": "FFFM", "name": "Foxglove Mix", "cat": "Flowers", "img": "https://i.ibb.co/23HXfz58/Foxglove-Foxy-Mix.png", "stock": 100}, {"id": 71, "code": "FPPu", "name": "Poppy Purple", "cat": "Flowers", "img": "https://i.ibb.co/p6srV4F5/Poppy-Purple.png", "stock": 100}, {"id": 72, "code": "FFLa", "name": "French Lavender", "cat": "Flowers", "img": "https://i.ibb.co/8L1QjnBn/Lavender-French.png", "stock": 100}, {"id": 73, "code": "FPWh", "name": "Poppy White", "cat": "Flowers", "img": "https://i.ibb.co/RTjmPqBQ/Poppy-White.png", "stock": 100}, {"id": 74, "code": "FHSa", "name": "Hibiscus Sabdariffa", "cat": "Flowers", "img": "https://i.ibb.co/b59sMBtf/Hibiscus-Sabdariffa-Roselle.png", "stock": 100}, {"id": 75, "code": "FPor", "name": "Portulaca", "cat": "Flowers", "img": "https://i.ibb.co/xrVyKrM/Portulaca-Mix.png", "stock": 100}, {"id": 76, "code": "FHMa", "name": "Hollyhock Maroon", "cat": "Flowers", "img": "https://i.ibb.co/HTCcpn4P/Hollyhock-Maroon.png", "stock": 100}, {"id": 77, "code": "FRud", "name": "Rudbeckia", "cat": "Flowers", "img": "https://i.ibb.co/4Zkv6xGF/Rudbeckia.png", "stock": 100}, {"id": 78, "code": "FHMi", "name": "Hollyhock Mix", "cat": "Flowers", "img": "https://i.ibb.co/YB1KpJ5C/Hollyhock-Mix.png", "stock": 100}, {"id": 79, "code": "FSTT", "name": "Snapdragon Tom Thumb Mix", "cat": "Flowers", "img": "https://i.ibb.co/Y7BcQtqf/Snapdragon-Tom-Thumb-Mix.png", "stock": 100}, {"id": 80, "code": "FHWh", "name": "Hollyhock White", "cat": "Flowers", "img": "https://i.ibb.co/TBGqwTJy/Hollyhock-White.png", "stock": 100}, {"id": 81, "code": "FSNi", "name": "Strelitzia Nicolai", "cat": "Flowers", "img": "https://i.ibb.co/qX03DMJ/Strelitzia-Nicoli.png", "stock": 100}, {"id": 82, "code": "FKPr", "name": "King Protea", "cat": "Flowers", "img": "https://i.ibb.co/3mH19z9K/Protea-King.png", "stock": 50}, {"id": 83, "code": "FSRe", "name": "Strelitzia Reginae", "cat": "Flowers", "img": "https://i.ibb.co/39Q2fXvG/Strelitzia-Regina.png", "stock": 0}, {"id": 84, "code": "FSMM", "name": "Summer Meadow Mix", "cat": "Flowers", "img": "https://i.ibb.co/NhnrHKx/Wildflower-Summer-Meadows-Mix.png", "stock": 100}, {"id": 85, "code": "FSWM", "name": "Sweet William Mix", "cat": "Flowers", "img": "https://i.ibb.co/jvCC8crk/Sweet-William-Mix.png", "stock": 100}, {"id": 86, "code": "FSBA", "name": "Sunflower Burnt Amber", "cat": "Flowers", "img": "https://i.ibb.co/7xKLW1Tv/Sunflower-Burnt-Amber.png", "stock": 50}, {"id": 87, "code": "FTul", "name": "Tulbaghia", "cat": "Flowers", "img": "https://i.ibb.co/YJmdzv7/Tulbaghia.png", "stock": 100}, {"id": 88, "code": "FSES", "name": "Sunflower Evening Star", "cat": "Flowers", "img": "https://i.ibb.co/VWtnfqZV/Sunflower-Evening-Star.png", "stock": 50}, {"id": 89, "code": "FVio", "name": "Viola", "cat": "Flowers", "img": "https://i.ibb.co/ynsPzxyN/Viola.png", "stock": 50}, {"id": 90, "code": "FSMi", "name": "Sunflower Mixed", "cat": "Flowers", "img": "https://i.ibb.co/Sw4mhsST/Sunflower-Mix.png", "stock": 50}, {"id": 91, "code": "FWMM", "name": "Winter Meadow Mix", "cat": "Flowers", "img": "https://i.ibb.co/5xg20csY/wildflower-Winter-Meadow-Mix.png", "stock": 100}, {"id": 92, "code": "FSNO", "name": "Sunflower Nigerian Oil", "cat": "Flowers", "img": "https://i.ibb.co/LDqGvfdF/Sunflower-Nigerian-Oil.png", "stock": 50}, {"id": 93, "code": "FYPP", "name": "Yellow Pinchushion Protea", "cat": "Flowers", "img": "https://i.ibb.co/Tx8jk7wk/Protea-Yellow-Pincushion.png", "stock": 100}, {"id": 94, "code": "FSTa", "name": "Sunflower Tarahumara", "cat": "Flowers", "img": "https://i.ibb.co/G4SwDqYQ/Sunflower-Tarahumara.png", "stock": 50}, {"id": 95, "code": "FZMi", "name": "Zinnia Mix", "cat": "Flowers", "img": "https://i.ibb.co/cSQc0P9V/Zinnia-Mix.png", "stock": 100}, {"id": 96, "code": "FSTE", "name": "Sunflower Tiger Eye", "cat": "Flowers", "img": "https://i.ibb.co/Fb5x5PpM/Sunflower-Tiger-Eye.png", "stock": 50}, {"id": 97, "code": "FSTi", "name": "Sunflower Titan", "cat": "Flowers", "img": "https://i.ibb.co/fd48jsZ6/Sunflower-Titan.png", "stock": 50}, {"id": 98, "code": "FCGo", "name": "Cape Gooseberry", "cat": "Fruit", "img": "https://i.ibb.co/mCdkdQRJ/Cape-Gooseberry.png", "stock": 100}, {"id": 99, "code": "FWAS", "name": "Watermelon All Sweet", "cat": "Fruit", "img": "https://i.ibb.co/fVYfRCNb/Watermelon-All-Sweet.png", "stock": 50}, {"id": 100, "code": "FGPP", "name": "Green Prickly Pear", "cat": "Fruit", "img": "https://i.ibb.co/YT0gnf2H/Green-Prickly-Pear.png", "stock": 100}, {"id": 101, "code": "FWBD", "name": "Watermelon Black Diamond", "cat": "Fruit", "img": "https://i.ibb.co/GfPHnZWs/Watermelon-Black-Diamond.png", "stock": 50}, {"id": 102, "code": "FHuc", "name": "Huckleberry", "cat": "Fruit", "img": "https://i.ibb.co/pB9vjJYN/Huckleberry.png", "stock": 50}, {"id": 103, "code": "FWGH", "name": "Watermelon Golden Honey", "cat": "Fruit", "img": "https://i.ibb.co/bRKmM75s/Watermelon-Golden-Honey.png", "stock": 50}, {"id": 104, "code": "FMHG", "name": "Melon Honeydew Green", "cat": "Fruit", "img": "https://i.ibb.co/nswLFSKk/Melon-Honeydew-Green.png", "stock": 50}, {"id": 105, "code": "FWMa", "name": "Watermelon Maketaan", "cat": "Fruit", "img": "https://i.ibb.co/cKWct0yP/Watermelon-Maketaan.png", "stock": 50}, {"id": 106, "code": "FMHY", "name": "Melon Honeydew Yellow", "cat": "Fruit", "img": "https://i.ibb.co/FGv3SzQ/Melon-Honeydew-Yellow.png", "stock": 50}, {"id": 107, "code": "FWMi", "name": "Watermelon Mix", "cat": "Fruit", "img": "https://i.ibb.co/Y40WgDV9/Watermelon-Mix.png", "stock": 50}, {"id": 108, "code": "FMMM", "name": "Melon Minnesota Midget", "cat": "Fruit", "img": "https://i.ibb.co/wNxWsQGD/Melon-Minnesota-Midget.png", "stock": 50}, {"id": 109, "code": "FWTS", "name": "Watermelon Tender Sweet Orange", "cat": "Fruit", "img": "https://i.ibb.co/ymQDVnVj/Watermelon-Tender-Sweet-Orange.png", "stock": 50}, {"id": 110, "code": "FMPD", "name": "Melon Piel De Sapo", "cat": "Fruit", "img": "https://i.ibb.co/VYJHgLXt/Melon-Piel-De-Sapo.jpg", "stock": 50}, {"id": 111, "code": "FMRS", "name": "Melon Rich Sweetness", "cat": "Fruit", "img": "https://i.ibb.co/rfTRNKhf/Melon-Rich-Sweetness-132.png", "stock": 50}, {"id": 112, "code": "GGAC", "name": "Gourd African Calabash", "cat": "Gourds", "img": "https://i.ibb.co/Fk0BnQgx/Gourd-African-Calabash.png", "stock": 30}, {"id": 113, "code": "GGGB", "name": "Gourd Giant Bullet Headwax", "cat": "Gourds", "img": "https://i.ibb.co/spqy9NNB/Gourd-Giant-Bullet-Headwax.png", "stock": 30}, {"id": 114, "code": "GGCB", "name": "Gourd Calabash Birdhouse", "cat": "Gourds", "img": "https://i.ibb.co/93Zynp8P/Gourd-Calabash-Birdhouse.png", "stock": 30}, {"id": 115, "code": "GGLe", "name": "Gourd Lerka", "cat": "Gourds", "img": "https://i.ibb.co/sd9PyC0Z/Gourd-Lerka.png", "stock": 30}, {"id": 116, "code": "GGCC", "name": "Gourd Caveman Club", "cat": "Gourds", "img": "https://i.ibb.co/23KSr7fz/Gourd-Caveman-Club.png", "stock": 30}, {"id": 117, "code": "GGMH", "name": "Gourd Martin House Bottle", "cat": "Gourds", "img": "https://i.ibb.co/HLNfNZvf/Gourd-Martin-House-Bottle.png", "stock": 30}, {"id": 118, "code": "GGCu", "name": "Gourd Cuzzuza", "cat": "Gourds", "img": "https://i.ibb.co/kV0KQz9h/Gourd-Cuzzuza.png", "stock": 30}, {"id": 119, "code": "GGMi", "name": "Gourd Mix", "cat": "Gourds", "img": "https://i.ibb.co/99Nhg1hM/Gourd-Mix.png", "stock": 30}, {"id": 120, "code": "GGDi", "name": "Gourd Dipper", "cat": "Gourds", "img": "https://i.ibb.co/7dYH92Cm/Gourd-Dipper.png", "stock": 30}, {"id": 121, "code": "GGSn", "name": "Gourd Snake", "cat": "Gourds", "img": "https://i.ibb.co/1GwzTjMX/Gourd-Snake.png", "stock": 30}, {"id": 122, "code": "GGDE", "name": "Gourd Dipper Extra Length", "cat": "Gourds", "img": "https://i.ibb.co/kZJ6Sd7/Gourd-Dipper-Extra-Length.png", "stock": 30}, {"id": 123, "code": "GGSS", "name": "Gourd Speckled Swan", "cat": "Gourds", "img": "https://i.ibb.co/39Kqgx65/Gourd-Speckled-Swan.png", "stock": 30}, {"id": 124, "code": "HAni", "name": "Anise", "cat": "Herbs", "img": "https://i.ibb.co/Ng6Xghqb/Anise.png", "stock": 100}, {"id": 125, "code": "HFev", "name": "Feverfew", "cat": "Herbs", "img": "https://i.ibb.co/8nvXn98S/Feverfew.png", "stock": 100}, {"id": 126, "code": "HAsh", "name": "Ashwagandha", "cat": "Herbs", "img": "https://i.ibb.co/7NRHb22B/Ashwagandha.png", "stock": 50}, {"id": 127, "code": "HFla", "name": "Flaxseed", "cat": "Herbs", "img": "https://i.ibb.co/TM88YRJx/Flaxseed.png", "stock": 100}, {"id": 128, "code": "HBGe", "name": "Basil Genovese", "cat": "Herbs", "img": "https://i.ibb.co/nT4vcKq/Basil-Genovese.png", "stock": 100}, {"id": 129, "code": "HHWh", "name": "Horehound White", "cat": "Herbs", "img": "https://i.ibb.co/svJdVJX5/Horehound-White.png", "stock": 100}, {"id": 130, "code": "HBHT", "name": "Basil Holy Tulsi", "cat": "Herbs", "img": "https://i.ibb.co/kgT2j4dW/Basil-Holy-Tulsi.png", "stock": 100}, {"id": 131, "code": "HHys", "name": "Hyssop", "cat": "Herbs", "img": "https://i.ibb.co/qYSM0mzQ/Hyssop.png", "stock": 100}, {"id": 132, "code": "HBLe", "name": "Basil Lemon", "cat": "Herbs", "img": "https://i.ibb.co/mVgtCytb/Basil-Lemon.png", "stock": 100}, {"id": 133, "code": "HJJT", "name": "Jacob's/Job's Tears (Coix Lacryma)", "cat": "Herbs", "img": "https://i.ibb.co/whx0D9cG/Jacob-Tears.png", "stock": 100}, {"id": 134, "code": "HBSw", "name": "Basil Sweet", "cat": "Herbs", "img": "https://i.ibb.co/1fbVBdh6/Basil-Sweet.png", "stock": 100}, {"id": 135, "code": "HLBa", "name": "Lemon Balm", "cat": "Herbs", "img": "https://i.ibb.co/ycNSkcfm/Lemon-Balm.png", "stock": 100}, {"id": 136, "code": "HBMu", "name": "Black Mustard", "cat": "Herbs", "img": "https://i.ibb.co/W4wd8CdY/Black-Mustard.png", "stock": 100}, {"id": 137, "code": "HLov", "name": "Lovage", "cat": "Herbs", "img": "https://i.ibb.co/JwcPbq09/Lovage.png", "stock": 100}, {"id": 138, "code": "HBSN", "name": "Black Seed (Nigella Sativa)", "cat": "Herbs", "img": "https://i.ibb.co/8DkFpBtG/Black-Seed-Nigella-Sativa.png", "stock": 100}, {"id": 139, "code": "HMMi", "name": "Mexican Mint (Plectranthus Amboinicus)", "cat": "Herbs", "img": "https://i.ibb.co/FdXfsXB/Mexican-Mint.png", "stock": 100}, {"id": 140, "code": "HBor", "name": "Borage", "cat": "Herbs", "img": "https://i.ibb.co/pjQP5hbk/Borage.png", "stock": 100}, {"id": 141, "code": "HMTh", "name": "Milk Thistle (Silybum Marianum)", "cat": "Herbs", "img": "https://i.ibb.co/Jg0gzKS/Milk-Thistle.png", "stock": 100}, {"id": 142, "code": "HCBu", "name": "Cancer Bush (Sutherlandia Frutescens)", "cat": "Herbs", "img": "https://i.ibb.co/x9SvRQ2/Cancer-Bush-Sutherlandia-Frustescens.png", "stock": 100}, {"id": 143, "code": "HMor", "name": "Moringa", "cat": "Herbs", "img": "https://i.ibb.co/Yg5PQBq/Moringa.png", "stock": 100}, {"id": 144, "code": "HCat", "name": "Catnip", "cat": "Herbs", "img": "https://i.ibb.co/dJBgpzk5/Catnip.png", "stock": 100}, {"id": 145, "code": "HMBe", "name": "Mung Beans", "cat": "Herbs", "img": "https://i.ibb.co/wZ2BVMGm/Mung-Bean.png", "stock": 100}, {"id": 146, "code": "HCha", "name": "Chamomile", "cat": "Herbs", "img": "https://i.ibb.co/hFr8n3s8/Chamomile.png", "stock": 100}, {"id": 147, "code": "HOre", "name": "Oregano", "cat": "Herbs", "img": "https://i.ibb.co/Y4GW2bff/Oregano.png", "stock": 100}, {"id": 148, "code": "HCSe", "name": "Chia Seed", "cat": "Herbs", "img": "https://i.ibb.co/gLwPwsqr/Chia-Seed.png", "stock": 100}, {"id": 149, "code": "HPFL", "name": "Parsley Flat Leaf", "cat": "Herbs", "img": "https://i.ibb.co/8LnqcTFG/Parsley-Flat-Leaf.png", "stock": 100}, {"id": 150, "code": "HChi", "name": "Chives", "cat": "Herbs", "img": "https://i.ibb.co/zWkPQJpt/Chives.png", "stock": 100}, {"id": 151, "code": "HPMC", "name": "Parsley Moss Curled", "cat": "Herbs", "img": "https://i.ibb.co/hxtc7DfQ/Parsley-Moss-Curled.png", "stock": 100}, {"id": 152, "code": "HCSa", "name": "Clarey Sage", "cat": "Herbs", "img": "https://i.ibb.co/HpMDj4Hv/Clarey-Sage.png", "stock": 100}, {"id": 153, "code": "HPur", "name": "Purslane", "cat": "Herbs", "img": "https://i.ibb.co/rRnSc5yx/Purslane.png", "stock": 0}, {"id": 154, "code": "HC/B", "name": "Cleaver/Bedstraw", "cat": "Herbs", "img": "https://i.ibb.co/yc9pkMPv/Cleaver.png", "stock": 100}, {"id": 155, "code": "HRMG", "name": "Red Mustard Giant Greens", "cat": "Herbs", "img": "https://i.ibb.co/PG8kRB9C/Red-Mustard-Giant-Greens.png", "stock": 100}, {"id": 156, "code": "HCor", "name": "Coriander", "cat": "Herbs", "img": "https://i.ibb.co/zWXh9wxs/Coriander.png", "stock": 100}, {"id": 157, "code": "HRAr", "name": "Rocket Arugula", "cat": "Herbs", "img": "https://i.ibb.co/B2pkQNbG/Rocket-Arugula.png", "stock": 100}, {"id": 158, "code": "HCTh", "name": "Creeping Thyme", "cat": "Herbs", "img": "https://i.ibb.co/TxgY4WCR/Creeping-Thyme.png", "stock": 100}, {"id": 159, "code": "HRSy", "name": "Rocket Sylvetta", "cat": "Herbs", "img": "https://i.ibb.co/zK0DTns/Rocket-Sylvetta.png", "stock": 100}, {"id": 160, "code": "HDan", "name": "Dandelion", "cat": "Herbs", "img": "https://i.ibb.co/4RB2Nbpf/Dandelion.png", "stock": 100}, {"id": 161, "code": "HRue", "name": "Rue", "cat": "Herbs", "img": "https://i.ibb.co/TMnKrBCk/Rue-Wynruit.png", "stock": 100}, {"id": 162, "code": "HDil", "name": "Dill", "cat": "Herbs", "img": "https://i.ibb.co/rGtdkJgk/Dill.png", "stock": 100}, {"id": 163, "code": "HSBu", "name": "Salad Burnet", "cat": "Herbs", "img": "https://i.ibb.co/3mYsL2xQ/Salad-Burnett.png", "stock": 100}, {"id": 164, "code": "HEch", "name": "Echinacea", "cat": "Herbs", "img": "https://i.ibb.co/R49VBVxC/Echinacea.png", "stock": 100}, {"id": 165, "code": "HSOn", "name": "Spring Onion", "cat": "Herbs", "img": "https://i.ibb.co/WTjpMCh/Spring-Onion.png", "stock": 100}, {"id": 166, "code": "HEPr", "name": "Evening Primrose", "cat": "Herbs", "img": "https://i.ibb.co/wNqsxWLb/Evening-Primrose.png", "stock": 100}, {"id": 167, "code": "HSNe", "name": "Stinging Nettle", "cat": "Herbs", "img": "https://i.ibb.co/23HNj9FL/Stinging-Nettle.png", "stock": 100}, {"id": 168, "code": "HFenn", "name": "Fennel", "cat": "Herbs", "img": "https://i.ibb.co/jPPfn7HK/Fennel.png", "stock": 100}, {"id": 169, "code": "HSun", "name": "Sunhemp", "cat": "Herbs", "img": "https://i.ibb.co/35hsZxLX/Sunhemp.png", "stock": 100}, {"id": 170, "code": "HFenu", "name": "Fenugreek", "cat": "Herbs", "img": "https://i.ibb.co/TDJS1PsY/Fenugreek.png", "stock": 100}, {"id": 171, "code": "HTar", "name": "Tarragon", "cat": "Herbs", "img": "https://i.ibb.co/nsDjbGwW/Tarragon.png", "stock": 100}, {"id": 172, "code": "HThy", "name": "Thyme", "cat": "Herbs", "img": "https://i.ibb.co/yLNrVg3/Thyme.png", "stock": 100}, {"id": 173, "code": "HTPl", "name": "Toothache Plant", "cat": "Herbs", "img": "https://i.ibb.co/zh8SPsM8/Toothache-Plant.png", "stock": 100}, {"id": 174, "code": "MMAl", "name": "Microgreen Alfalfa", "cat": "Microgreens", "img": "https://i.ibb.co/2Yyn64Vf/Microgreen-Alfalfa.png", "stock": 50}, {"id": 175, "code": "MMRA", "name": "Microgreen Red Amaranth", "cat": "Microgreens", "img": "https://i.ibb.co/GQsDfMh4/Microgreen-Red-Amaranth.png", "stock": 50}, {"id": 176, "code": "MMAO", "name": "Microgreen Asian Oriental Blend", "cat": "Microgreens", "img": "https://i.ibb.co/chH8CTch/Microgreen-Asian-Oriental-Blend.png", "stock": 50}, {"id": 177, "code": "MMRS", "name": "Microgreen Red Swiss Chard", "cat": "Microgreens", "img": "https://i.ibb.co/nN3tVfcs/Microgreen-Red-Swiss-Chard.png", "stock": 50}, {"id": 178, "code": "MMBM", "name": "Microgreen Black Mustard", "cat": "Microgreens", "img": "https://i.ibb.co/62QmzSw/Microgreen-Black-Mustard.jpg", "stock": 50}, {"id": 179, "code": "MMRo", "name": "Microgreen Rocket", "cat": "Microgreens", "img": "https://i.ibb.co/wNgwZBWG/Microgreen-Rocket.png", "stock": 50}, {"id": 180, "code": "MMBC", "name": "Microgreen Broccoli Calabrese", "cat": "Microgreens", "img": "https://i.ibb.co/tTHDN9QM/Microgreen-Broccoli.png", "stock": 50}, {"id": 181, "code": "MMSB", "name": "Microgreen Stirfry Blend", "cat": "Microgreens", "img": "https://i.ibb.co/PyGC3qY/Microgreen-Stirfry-Blend.png", "stock": 50}, {"id": 182, "code": "MMCr", "name": "Microgreen Cress", "cat": "Microgreens", "img": "https://i.ibb.co/d0h7jJxL/Microgreen-Cress.png", "stock": 50}, {"id": 183, "code": "MMSS", "name": "Microgreen Striped Sunflower", "cat": "Microgreens", "img": "https://i.ibb.co/8n1Qb9bn/Microgreen-Striped-Sunflowers.png", "stock": 50}, {"id": 184, "code": "MMFe", "name": "Microgreen Fenugreek", "cat": "Microgreens", "img": "https://i.ibb.co/rGG5qNhX/Microgreen-Fenugreek.jpg", "stock": 50}, {"id": 185, "code": "MMWh", "name": "Microgreen Wheatgrass", "cat": "Microgreens", "img": "https://i.ibb.co/bRrQMbFR/Microgreen-Wheatgrass.png", "stock": 50}, {"id": 186, "code": "MMGB", "name": "Microgreen Green Basil", "cat": "Microgreens", "img": "https://i.ibb.co/BHW0pNxF/Microgreen-Green-Basil-5.jpg", "stock": 50}, {"id": 187, "code": "MMWM", "name": "Microgreen White Mustard", "cat": "Microgreens", "img": "https://i.ibb.co/JwNChQ6K/Microgreen-White-Mustard.png", "stock": 50}, {"id": 188, "code": "MMGS", "name": "Microgreen Green Swiss Chard", "cat": "Microgreens", "img": "https://i.ibb.co/wtBMrh5/Microgreen-Green-Swiss-Chard.jpg", "stock": 50}, {"id": 189, "code": "MMYM", "name": "Microgreen Yellow Mustard", "cat": "Microgreens", "img": "https://i.ibb.co/VYFMvXSH/Microgreen-Yellow-Mustard.png", "stock": 50}, {"id": 190, "code": "MMKM", "name": "Microgreen Kaleidoscope Mix", "cat": "Microgreens", "img": "https://i.ibb.co/kgfyxhnQ/Microgreen-Kaleidoscope.png", "stock": 50}, {"id": 191, "code": "MMMGr", "name": "Microgreen Mizuna Greens", "cat": "Microgreens", "img": "https://i.ibb.co/TM39QVQf/Microgreen-Mizuno-Green.png", "stock": 50}, {"id": 192, "code": "MMMu", "name": "Microgreen Mungbeans", "cat": "Microgreens", "img": "https://i.ibb.co/GQZKX5Cv/Microgreen-Mungbeans.png", "stock": 50}, {"id": 193, "code": "MMMGf", "name": "Microgreen Mustard Green Frills", "cat": "Microgreens", "img": "https://i.ibb.co/RpHd7f0X/Microgreen-Mustard-Greens-Frills.png", "stock": 50}, {"id": 194, "code": "MMMRf", "name": "Microgreen Mustard Red Frills", "cat": "Microgreens", "img": "https://i.ibb.co/z938tfR/Microgreen-Mustard-Red-Frills.png", "stock": 50}, {"id": 195, "code": "MMRC", "name": "Microgreen Radish Coralette", "cat": "Microgreens", "img": "https://i.ibb.co/qLSCzM7M/Microgreen-Radish-Coralette.png", "stock": 50}, {"id": 196, "code": "MMRD", "name": "Microgreen Radish Daikon", "cat": "Microgreens", "img": "https://i.ibb.co/p87RfY2/Microgreen-Radish-Daikon.png", "stock": 50}, {"id": 197, "code": "MMRR", "name": "Microgreen Radish Rainbow", "cat": "Microgreens", "img": "https://i.ibb.co/byCVjxg/Microgreen-Radish-Rainbow-Mix.png", "stock": 50}, {"id": 198, "code": "MMRT", "name": "Microgreen Radish Tsai Tsai", "cat": "Microgreens", "img": "https://i.ibb.co/ccjrtNNd/Microgreen-Radish-Tsai-Tsai.png", "stock": 50}, {"id": 199, "code": "PPBG", "name": "Peanut Behin Giant Striped", "cat": "Peanuts", "img": "https://i.ibb.co/5XqdsPS5/Peanut-Benih-Giant-Strip.png", "stock": 30}, {"id": 200, "code": "PPBl", "name": "Peanut Black", "cat": "Peanuts", "img": "https://i.ibb.co/RkjfLszQ/Peanut-Black.png", "stock": 30}, {"id": 201, "code": "PPCh", "name": "Peanut Chalimbana", "cat": "Peanuts", "img": "https://i.ibb.co/JwWsxnpd/Peanut-Chalimbana.png", "stock": 30}, {"id": 202, "code": "PPFP", "name": "Peanut Fastigianta Pin Striped", "cat": "Peanuts", "img": "https://i.ibb.co/LMh7d6c/Peanut-Fastigianta-Pin-Stripped.png", "stock": 30}, {"id": 203, "code": "PPMS", "name": "Peanut Malawi Striped", "cat": "Peanuts", "img": "https://i.ibb.co/gFtYRYkN/Peanut-Malawi-Stripe.png", "stock": 30}, {"id": 204, "code": "SSML", "name": "Sprout Mansoor Lentils", "cat": "Sprouts", "img": "https://i.ibb.co/pBj3ZNL1/Sprouts-Mansoor-Lentils.png", "stock": 50}, {"id": 205, "code": "SSMB", "name": "Sprout Moth Beans", "cat": "Sprouts", "img": "https://i.ibb.co/YBZtYr8k/Sprouts-Moth-Beans.png", "stock": 50}, {"id": 206, "code": "SSMu", "name": "Sprout Mungbeans", "cat": "Sprouts", "img": "https://i.ibb.co/d4HwQBDy/Sprouts-Mungbeans.png", "stock": 50}, {"id": 207, "code": "SSSP", "name": "Sprout Sprouting Peas", "cat": "Sprouts", "img": "https://i.ibb.co/qMQpbPh6/Sprouts-Sprouting-Peas.png", "stock": 50}, {"id": 208, "code": "SSSB", "name": "Sprout Stirfry Blend", "cat": "Sprouts", "img": "https://i.ibb.co/MDCWPKgH/Sprouts-Stirfry-Blend.png", "stock": 50}, {"id": 209, "code": "SSSS", "name": "Sprout Sunflower Seed", "cat": "Sprouts", "img": "https://i.ibb.co/xqL4swx8/Sprouts-Sunflower.png", "stock": 50}, {"id": 210, "code": "SSTa", "name": "Sprout Tatsoi", "cat": "Sprouts", "img": "https://i.ibb.co/pjR9L013/Sprouts-Tatsoi.png", "stock": 50}, {"id": 211, "code": "SSTK", "name": "Sprout Tuscany Kale", "cat": "Sprouts", "img": "https://i.ibb.co/CsXFb1rN/Sprouts-Tuscany-Kale.png", "stock": 50}, {"id": 212, "code": "SSWC", "name": "Sprout White Chickpeas", "cat": "Sprouts", "img": "https://i.ibb.co/h1RFmbsJ/Sprouts-White-Chickpeas.png", "stock": 50}, {"id": 213, "code": "TBao", "name": "Baobab", "cat": "Trees", "img": "https://i.ibb.co/tptsYJ3P/Boabab.png", "stock": 0}, {"id": 214, "code": "TSKT", "name": "Swart Karee Tree", "cat": "Trees", "img": "https://i.ibb.co/7NJLSwvC/Swart-Karee-Tree.png", "stock": 100}, {"id": 215, "code": "TCTT", "name": "Camel Thorn Tree", "cat": "Trees", "img": "https://i.ibb.co/tpZNjhvb/Camel-Thorn-Tree.png", "stock": 100}, {"id": 216, "code": "TSTT", "name": "Sweet Thorn Tree", "cat": "Trees", "img": "https://i.ibb.co/jkDVTvy9/Sweet-Thorn-Tree.png", "stock": 100}, {"id": 217, "code": "TCTr", "name": "Carol Tree", "cat": "Trees", "img": "https://i.ibb.co/wFdL8jLh/Coral-Tree.png", "stock": 0}, {"id": 218, "code": "TWOT", "name": "Wild Olive Tree", "cat": "Trees", "img": "https://i.ibb.co/DPdFhHmD/Wild-Olive-Tree.png", "stock": 0}, {"id": 219, "code": "TFTr", "name": "Fever Tree", "cat": "Trees", "img": "https://i.ibb.co/DHMzt529/Fever-Tree.jpg", "stock": 100}, {"id": 220, "code": "TLTr", "name": "Lucern Tree", "cat": "Trees", "img": "https://i.ibb.co/yFDXK0n4/Lucerne-Tree.jpg", "stock": 100}, {"id": 221, "code": "TMTr", "name": "Marula Tree", "cat": "Trees", "img": "https://i.ibb.co/Y7Npf7zV/Marula-Tree.png", "stock": 100}, {"id": 222, "code": "TMTT", "name": "Monkey Thorn Tree", "cat": "Trees", "img": "https://i.ibb.co/LzgDYbkM/Monkey-Thorn-Tree.jpg", "stock": 100}, {"id": 223, "code": "TPTT", "name": "Paperback Thorn Tree", "cat": "Trees", "img": "https://i.ibb.co/QvD5J8Sr/Paper-Bark-Thorn-Tree.jpg", "stock": 100}, {"id": 224, "code": "TSOT", "name": "Sand Olive Tree", "cat": "Trees", "img": "https://i.ibb.co/hx0VDmCf/Sand-Olive-Tree.png", "stock": 100}, {"id": 225, "code": "VAHC", "name": "African Horned Cucumber", "cat": "Vegetables", "img": "https://i.ibb.co/rqnzxsx/African-Horned-Cucumber.png", "stock": 100}, {"id": 226, "code": "VCLe", "name": "Cucumber Lemon", "cat": "Vegetables", "img": "https://i.ibb.co/WWbyxbmy/Cucumber-Lemon.png", "stock": 100}, {"id": 227, "code": "VAGG", "name": "Artichoke Green Globe", "cat": "Vegetables", "img": "https://i.ibb.co/MDVJ4LjD/Artichoke-Green-Globe.png", "stock": 100}, {"id": 228, "code": "VCLW", "name": "Cucumber Long White", "cat": "Vegetables", "img": "https://i.ibb.co/0p3KpGSq/Cucmber-Long-White.png", "stock": 100}, {"id": 229, "code": "VAMW", "name": "Asparagus Mary Washington", "cat": "Vegetables", "img": "https://i.ibb.co/Q7zS0GRN/Asparagus-Mary-Washington.png", "stock": 100}, {"id": 230, "code": "VCSnk", "name": "Cucumber Snake", "cat": "Vegetables", "img": "https://i.ibb.co/cGcmBCL/Cucumber-Snake.png", "stock": 100}, {"id": 231, "code": "VBAB", "name": "Bean Appaloosa Bush", "cat": "Vegetables", "img": "https://i.ibb.co/Hppqf9yK/Bean-Appaloosa-Bush.png", "stock": 100}, {"id": 232, "code": "VGCR", "name": "Garlic Chesnok Red (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/qMhvN3MV/Garlic-Chesnok-Red.jpg", "stock": 0}, {"id": 233, "code": "VBBE", "name": "Bean Bird Egg Blue", "cat": "Vegetables", "img": "https://i.ibb.co/WNz347bV/Bean-Bird-Egg-Blue.png", "stock": 100}, {"id": 234, "code": "VGER", "name": "Garlic Egyptian Red (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/fz70qDpk/Garlic-Egyptian-Red.jpg", "stock": 0}, {"id": 235, "code": "VBBT", "name": "Bean Black Turtle", "cat": "Vegetables", "img": "https://i.ibb.co/hxZcpc7V/Bean-Black-Turtle.png", "stock": 100}, {"id": 236, "code": "VGEW", "name": "Garlic Egyptain White (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/V0R9CHCj/Garlic-Egypitain-White.jpg", "stock": 0}, {"id": 237, "code": "VBBrb", "name": "Bean Broad Bean", "cat": "Vegetables", "img": "https://i.ibb.co/jkdsBBqt/Bean-Broad-Bean.png", "stock": 100}, {"id": 238, "code": "VGSW", "name": "Garlic Spanish White (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/TDkWSVB7/Garlic-Spanish-White.jpg", "stock": 0}, {"id": 239, "code": "VBBC", "name": "Bean Bush Contender", "cat": "Vegetables", "img": "https://i.ibb.co/Nnb76zwB/Bean-Bush-Contender.png", "stock": 100}, {"id": 240, "code": "VKBP", "name": "Kale Black Palm", "cat": "Vegetables", "img": "https://i.ibb.co/4wrHS0TK/Kale-Black-Palm.png", "stock": 100}, {"id": 241, "code": "VBLN", "name": "Bean Lima Nuguni Pole", "cat": "Vegetables", "img": "https://i.ibb.co/kVz7VZ4B/Bean-Lima-Nuguni-Pole.png", "stock": 100}, {"id": 242, "code": "VKRU", "name": "Kale Red Ursa", "cat": "Vegetables", "img": "https://i.ibb.co/Kz92cB99/Kale-Red-Ursa.png", "stock": 100}, {"id": 243, "code": "VBMa", "name": "Bean Madagascar", "cat": "Vegetables", "img": "https://i.ibb.co/MxpNC5PV/Bean-Madagascar.png", "stock": 100}, {"id": 244, "code": "VKSB", "name": "Kale Southern Blue", "cat": "Vegetables", "img": "https://i.ibb.co/N6Jkk5pc/Kale-Southern-Blue.png", "stock": 100}, {"id": 245, "code": "VBNB", "name": "Bean Nonna Blue", "cat": "Vegetables", "img": "https://i.ibb.co/Q71pnbDf/Bean-Nonna-Blue-Agnes.png", "stock": 100}, {"id": 246, "code": "VKVB", "name": "Kale Vates Blue", "cat": "Vegetables", "img": "https://i.ibb.co/vxdTqsMg/Kale-Vates-Blue.png", "stock": 100}, {"id": 247, "code": "VBVA", "name": "Bean Vermont Appaloosa", "cat": "Vegetables", "img": "https://i.ibb.co/27v5hfXC/Bean-Vermont-Appaloosa.png", "stock": 100}, {"id": 248, "code": "VKWR", "name": "Kale White Russian", "cat": "Vegetables", "img": "https://i.ibb.co/Mkkd0QyL/Kale-White-Russain.png", "stock": 100}, {"id": 249, "code": "VBWi", "name": "Bean Witsa", "cat": "Vegetables", "img": "https://i.ibb.co/zhtTz2M0/Bean-Witsa.png", "stock": 100}, {"id": 250, "code": "VKPV", "name": "Kohlrabi Purple Vienna", "cat": "Vegetables", "img": "https://i.ibb.co/WN755px2/Kohlrabi-Purple-Vienna.png", "stock": 100}, {"id": 251, "code": "VBYL", "name": "Bean Yard Long", "cat": "Vegetables", "img": "https://i.ibb.co/HpR2nh2g/Bean-Yard-Long.png", "stock": 100}, {"id": 252, "code": "VKWV", "name": "Kohlrabi White Vienna", "cat": "Vegetables", "img": "https://i.ibb.co/gFWLX6ZZ/Kohlrabi-White-Vienna.png", "stock": 100}, {"id": 253, "code": "VBZL", "name": "Bean Zebra Lima", "cat": "Vegetables", "img": "https://i.ibb.co/b5ckSNyp/Bean-Zebra-Lima.png", "stock": 100}, {"id": 254, "code": "VLEl", "name": "Leeks Elephant", "cat": "Vegetables", "img": "https://i.ibb.co/KzjyyDhq/Leeks-Elephant.png", "stock": 100}, {"id": 255, "code": "VBBu", "name": "Beetroot Bulls Blood", "cat": "Vegetables", "img": "https://i.ibb.co/spR3GsJL/Beetroot-Bulls-Blood.png", "stock": 100}, {"id": 256, "code": "VLBu", "name": "Lettuce Butterhead", "cat": "Vegetables", "img": "https://i.ibb.co/DgMRCv0C/Lettuce-Butterhead.png", "stock": 100}, {"id": 257, "code": "VBCh", "name": "Beetroot Chioggia", "cat": "Vegetables", "img": "https://i.ibb.co/1GNHBL8v/Beetroot-Chiaggia.png", "stock": 100}, {"id": 258, "code": "VLGM", "name": "Lettuce Gourmet Mix", "cat": "Vegetables", "img": "https://i.ibb.co/xqHn6QrF/Lettuce-Gourmet-Mix.png", "stock": 100}, {"id": 259, "code": "VBCG", "name": "Beetroot Crimson Globe", "cat": "Vegetables", "img": "https://i.ibb.co/PvyqvC1f/Beetroot-Crimson-Globe.png", "stock": 100}, {"id": 260, "code": "VLGL", "name": "Lettuce Great Lakes", "cat": "Vegetables", "img": "https://i.ibb.co/RTqJkTCc/Lettuce-Great-Lakes.png", "stock": 100}, {"id": 261, "code": "VBDD", "name": "Beetroot Detroit Dark Red", "cat": "Vegetables", "img": "https://i.ibb.co/jvVHfKz6/Beetroot-Detroit-Dark-Red.png", "stock": 100}, {"id": 262, "code": "VLGC", "name": "Lettuce Green Cos", "cat": "Vegetables", "img": "https://i.ibb.co/KckGxDYZ/Lettuce-Green-Cos.png", "stock": 100}, {"id": 263, "code": "VBGG", "name": "Beetroot Golden Globe", "cat": "Vegetables", "img": "https://i.ibb.co/8n5Y63bB/Beetroot-Golden-Globe.png", "stock": 100}, {"id": 264, "code": "VLLL", "name": "Lettuce Loose Leaf Mix", "cat": "Vegetables", "img": "https://i.ibb.co/bMBC9sr2/Lettuce-Loose-Leaf-Mix.png", "stock": 100}, {"id": 265, "code": "VBRM", "name": "Beetroot Rainbow Mix", "cat": "Vegetables", "img": "https://i.ibb.co/Z6jnQhy8/Beetroot-Rainbow-Mix.png", "stock": 100}, {"id": 266, "code": "VLOL", "name": "Lettuce Oak Leaf Mix", "cat": "Vegetables", "img": "https://i.ibb.co/Jwf622m2/Lettuce-Oak-Lead-Mix.png", "stock": 100}, {"id": 267, "code": "VBRQ", "name": "Beetroot Ruby Queen", "cat": "Vegetables", "img": "https://i.ibb.co/whKyN6cn/Beetroot-Ruby-Queen.png", "stock": 100}, {"id": 268, "code": "VLRF", "name": "Lettuce Romaine Freckles", "cat": "Vegetables", "img": "https://i.ibb.co/Vc0GdhVG/Lettuce-Romaine-Freckles.png", "stock": 100}, {"id": 269, "code": "VBBi", "name": "Brinjal Bianca", "cat": "Vegetables", "img": "https://i.ibb.co/TXncMj3/Brinjal-Bianca.png", "stock": 100}, {"id": 270, "code": "VLRM", "name": "Lettuce Romain Mix", "cat": "Vegetables", "img": "https://i.ibb.co/8D48wm9X/Lettuce-Romaine-Mix.png", "stock": 100}, {"id": 271, "code": "VBBy", "name": "Brinjal Black Beauty", "cat": "Vegetables", "img": "https://i.ibb.co/006y90j/Brinjal-Black-Beauty.png", "stock": 100}, {"id": 272, "code": "VMGr", "name": "Marog Green", "cat": "Vegetables", "img": "https://i.ibb.co/WNtNXVhm/Marog-Green.png", "stock": 100}, {"id": 273, "code": "VBPF", "name": "Brinjal Purple Fingers", "cat": "Vegetables", "img": "https://i.ibb.co/bg4Vg9pq/Brinjal-Purple-Fingers.png", "stock": 100}, {"id": 274, "code": "VMSa", "name": "Mealie Sahara", "cat": "Vegetables", "img": "https://i.ibb.co/4w99Jtyf/Mealie-Sahara.png", "stock": 100}, {"id": 275, "code": "VBCa", "name": "Broccoli Calabrese", "cat": "Vegetables", "img": "https://i.ibb.co/vfBdCMj/Broccoli-Calabrese.png", "stock": 100}, {"id": 276, "code": "VMSp", "name": "Mustard Spinach", "cat": "Vegetables", "img": "https://i.ibb.co/1YMhvT8D/Mustard-Green-Spinach.png", "stock": 100}, {"id": 277, "code": "VBPS", "name": "Broccoli Purple Sprouting", "cat": "Vegetables", "img": "https://i.ibb.co/F4TB2pS7/Broccoli-Purple-Sprouting.png", "stock": 100}, {"id": 278, "code": "VOLF", "name": "Okra Lady Fingers", "cat": "Vegetables", "img": "https://i.ibb.co/J8h8dvS/Okra-Lady-Fingers.png", "stock": 100}, {"id": 279, "code": "VBRo", "name": "Broccoli Romanesca", "cat": "Vegetables", "img": "https://i.ibb.co/mrtj5Fh9/Broccoli-Romanessca.png", "stock": 100}, {"id": 280, "code": "VOAB", "name": "Onion Australian Brown", "cat": "Vegetables", "img": "https://i.ibb.co/NnghSvDT/Onion-Australian-Brown.png", "stock": 100}, {"id": 281, "code": "VBSL", "name": "Brussel Sprouts Long Island", "cat": "Vegetables", "img": "https://i.ibb.co/213Xmr4J/Brussels-Sprout.png", "stock": 100}, {"id": 282, "code": "VORC", "name": "Onion Red Creole", "cat": "Vegetables", "img": "https://i.ibb.co/p6sGnTzs/Onion-Red-Creole.png", "stock": 100}, {"id": 283, "code": "VCCo", "name": "Cabbage Copenhagen", "cat": "Vegetables", "img": "https://i.ibb.co/pvgGXYx6/Cabbage-Copenhagen.png", "stock": 100}, {"id": 284, "code": "VOTG", "name": "Onion Texas Grano", "cat": "Vegetables", "img": "https://i.ibb.co/VWPwKPZx/Onion-Texas-Grano.png", "stock": 100}, {"id": 285, "code": "VCDr", "name": "Cabbage Drumhead", "cat": "Vegetables", "img": "https://i.ibb.co/Psp3XQS8/Cabbage-Drumhead.png", "stock": 100}, {"id": 286, "code": "VOAV", "name": "Oriental Asian Vegetable Mix", "cat": "Vegetables", "img": "https://i.ibb.co/7xN7TVyM/Oriental-Asain-Blend-Mix.png", "stock": 100}, {"id": 287, "code": "VCSa", "name": "Cabbage Savoy", "cat": "Vegetables", "img": "https://i.ibb.co/hRLfHYsv/Cabbage-Savoy.png", "stock": 100}, {"id": 288, "code": "VPAr", "name": "Peas Aragon", "cat": "Vegetables", "img": "https://i.ibb.co/yFw8Cc1Q/Peas-Aargan.png", "stock": 100}, {"id": 289, "code": "VCCK", "name": "Carrot Chantenay Karoo", "cat": "Vegetables", "img": "https://i.ibb.co/TQ3GgSL/Carrot-Chantenay-Karoo.png", "stock": 100}, {"id": 290, "code": "VPFE", "name": "Peas First Early May", "cat": "Vegetables", "img": "https://i.ibb.co/ycF2wZns/Peas-First-Early-May.png", "stock": 100}, {"id": 291, "code": "VCNR", "name": "Carrot Nantes Red", "cat": "Vegetables", "img": "https://i.ibb.co/R4bGKQMx/Carrot-Nantes-Red.png", "stock": 100}, {"id": 292, "code": "VPSS", "name": "Peas Sugar Snap Mangetout", "cat": "Vegetables", "img": "https://i.ibb.co/HTLZ8crv/Peas-Sugar-Snap-Mange-Tout.png", "stock": 100}, {"id": 293, "code": "VCRM", "name": "Carrot Rainbow Mix", "cat": "Vegetables", "img": "https://i.ibb.co/JR6T4R77/Carrot-Rainbow-Mix.png", "stock": 100}, {"id": 294, "code": "VPSSn", "name": "Peas Super Snappy", "cat": "Vegetables", "img": "https://i.ibb.co/zT5wDw5f/Peas-Super-Snappy.png", "stock": 100}, {"id": 295, "code": "VCGI", "name": "Cauliflower Green Igloo", "cat": "Vegetables", "img": "https://i.ibb.co/YBPfwwFk/Cauliflower-Green-Igloo.png", "stock": 100}, {"id": 296, "code": "VPDB", "name": "Popcorn Dakota Black", "cat": "Vegetables", "img": "https://i.ibb.co/V0jgKmWP/Popcorn-Black-Dakota.png", "stock": 100}, {"id": 297, "code": "VCMG", "name": "Cauliflower Macerata Green", "cat": "Vegetables", "img": "https://i.ibb.co/WWNMP868/Cauliflower-Macerata-Green.png", "stock": 100}, {"id": 298, "code": "VPQB", "name": "Pumpkin Queensland Blue", "cat": "Vegetables", "img": "https://i.ibb.co/hJfGXXT5/Pumpkin-Queensland-Blue.png", "stock": 100}, {"id": 299, "code": "VCMW", "name": "Cauliflower Mini White", "cat": "Vegetables", "img": "https://i.ibb.co/MkgDMZNx/Cauliflower-Mini-White.png", "stock": 100}, {"id": 300, "code": "VPTT", "name": "Pumpkin Turks Turban", "cat": "Vegetables", "img": "https://i.ibb.co/bMjqNzgM/Pumpkin-Turks-Turban.png", "stock": 100}, {"id": 301, "code": "VCRG", "name": "Cauliflower Romanesca Green", "cat": "Vegetables", "img": "https://i.ibb.co/0jryjnv6/Cauliflower-Romanessca-Green.png", "stock": 100}, {"id": 302, "code": "VPWi", "name": "Pumpkin Witboer", "cat": "Vegetables", "img": "https://i.ibb.co/5gVS7Pm1/Pumpkin-Witboer.png", "stock": 100}, {"id": 303, "code": "VCSb", "name": "Cauliflower Snowball", "cat": "Vegetables", "img": "https://i.ibb.co/Xf1Fgq6w/Cauliflower-Snowball-White.png", "stock": 100}, {"id": 304, "code": "VRCB", "name": "Radish Cherry Belle", "cat": "Vegetables", "img": "https://i.ibb.co/RTWJkCmd/Radish-Cherry-Belle.png", "stock": 100}, {"id": 305, "code": "VCVS", "name": "Cauliflower Violet Sicilian", "cat": "Vegetables", "img": "https://i.ibb.co/k2cBFMNB/Cauliflower-Violet-Sicilian.png", "stock": 100}, {"id": 306, "code": "VRHW", "name": "Radish Hailstone White", "cat": "Vegetables", "img": "https://i.ibb.co/Rp7QjNbx/Radish-Hailstone-White.png", "stock": 100}, {"id": 307, "code": "VCUT", "name": "Celery Utah Tall", "cat": "Vegetables", "img": "https://i.ibb.co/HfTbjWHr/Celery-Utah-Tall.png", "stock": 100}, {"id": 308, "code": "VRPP", "name": "Radish Purple Plum", "cat": "Vegetables", "img": "https://i.ibb.co/SDrwx91c/Radish-Purple-Plum.png", "stock": 100}, {"id": 309, "code": "VCBB", "name": "Corn Bloody Butcher", "cat": "Vegetables", "img": "https://i.ibb.co/Kpr2zF7j/Corn-Bloody-Butcher.png", "stock": 100}, {"id": 310, "code": "VRRM", "name": "Radish Rainbow Mix", "cat": "Vegetables", "img": "https://i.ibb.co/1fJ5b2xN/Radish-Rainbow-Mix.png", "stock": 100}, {"id": 311, "code": "VCGG", "name": "Corn Glass Gem", "cat": "Vegetables", "img": "https://i.ibb.co/NggBDfTw/Corn-Glass-Gem.png", "stock": 100}, {"id": 312, "code": "VRSB", "name": "Radish Spanish Black", "cat": "Vegetables", "img": "https://i.ibb.co/ZpbV9Vmx/Radish-Spanish-Black.png", "stock": 100}, {"id": 313, "code": "VCGO", "name": "Corn Green Oaxacan", "cat": "Vegetables", "img": "https://i.ibb.co/tw9tpwN0/Corn-Green-Oaxacan.png", "stock": 100}, {"id": 314, "code": "VRSp", "name": "Radish Sparkler", "cat": "Vegetables", "img": "https://i.ibb.co/7d9XktwJ/Radish-Sparkler.png", "stock": 100}, {"id": 315, "code": "VCAW", "name": "Cucumber Armenian White", "cat": "Vegetables", "img": "https://i.ibb.co/Q79zs50g/Cucumber-Arminian-White.png", "stock": 100}, {"id": 316, "code": "VRWI", "name": "Radish White Icicle", "cat": "Vegetables", "img": "https://i.ibb.co/spk76h6j/Radish-White-Icicle.png", "stock": 100}, {"id": 317, "code": "VCAs", "name": "Cucumber Ashley", "cat": "Vegetables", "img": "https://i.ibb.co/bRJj0Wrb/Cucumber-Ashley.png", "stock": 100}, {"id": 318, "code": "VRSC", "name": "Red Swiss Chard", "cat": "Vegetables", "img": "https://i.ibb.co/nMJLQtx7/Red-Swiss-Chard.png", "stock": 100}, {"id": 319, "code": "VCDY", "name": "Cucumber Double Yield", "cat": "Vegetables", "img": "https://i.ibb.co/kVwv1dmX/Cucmber-Double-Yeild.png", "stock": 100}, {"id": 320, "code": "VRhu", "name": "Rhubarb", "cat": "Vegetables", "img": "https://i.ibb.co/hx1tmFKq/Rhubarb.png", "stock": 100}, {"id": 321, "code": "VCGR", "name": "Cucumber Gherkin Rhinish", "cat": "Vegetables", "img": "https://i.ibb.co/6VCHmgh/Cucumber-Gherkin-Rhinish.png", "stock": 100}, {"id": 322, "code": "VSBB", "name": "Spinach Baby Black Magic", "cat": "Vegetables", "img": "https://i.ibb.co/xqXtSk96/Spinach-Baby-Black-Magic.png", "stock": 100}, {"id": 323, "code": "VSAy", "name": "Squash Ayota", "cat": "Vegetables", "img": "https://i.ibb.co/Z6VpBKpR/Squash-Ayota.png", "stock": 0}, {"id": 324, "code": "VSGR", "name": "Squash Gem Rolet", "cat": "Vegetables", "img": "https://i.ibb.co/xtGzD0Rs/Squash-Gem-Rolet.png", "stock": 100}, {"id": 325, "code": "VSPA", "name": "Squash Patty Pan Alba", "cat": "Vegetables", "img": "https://i.ibb.co/1fM1xG1S/Squash-Patty-Pam-Alba-White.png", "stock": 100}, {"id": 326, "code": "VSPJ", "name": "Squah Patty Pan Juane", "cat": "Vegetables", "img": "https://i.ibb.co/QZYzYgv/Squash-Patty-Pan-Juane-Et-Verte.png", "stock": 100}, {"id": 327, "code": "VSPM", "name": "Squash Patty Pan Mix", "cat": "Vegetables", "img": "https://i.ibb.co/yBf7txWT/Squahs-Patty-Pan-Mix.png", "stock": 100}, {"id": 328, "code": "VSPY", "name": "Squash Patty Pan Yellow", "cat": "Vegetables", "img": "https://i.ibb.co/Z6Q0fDw3/Squash-Patty-Pan-Yellow.png", "stock": 100}, {"id": 329, "code": "VSSp", "name": "Squash Spaghetti", "cat": "Vegetables", "img": "https://i.ibb.co/Xkbtk4Yq/Squash-Spaghetti.png", "stock": 100}, {"id": 330, "code": "VSWB", "name": "Squash Waltham Butternut", "cat": "Vegetables", "img": "https://i.ibb.co/jPNPfdqL/Squash-Waltham-Butternut.png", "stock": 100}, {"id": 331, "code": "VSGB", "name": "Sweetcorn Golden Bantam", "cat": "Vegetables", "img": "https://i.ibb.co/Gr9bthG/Sweetcorn-Golden-Bantam.png", "stock": 100}, {"id": 332, "code": "VSPC", "name": "Sweet Potato Combo Pack (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/JRmMj0ZP/Sweet-Potato-Combo-Pack.jpg", "stock": 0}, {"id": 333, "code": "VSPOk", "name": "Sweet Potato Okinawan (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/j9X1H3gZ/Sweet-Potato-Okinawan.jpg", "stock": 0}, {"id": 334, "code": "VSPOr", "name": "Sweet Potato Orange (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/5g1FLkPk/Sweet-Potato-Orange.jpg", "stock": 0}, {"id": 335, "code": "VSPP", "name": "Sweet Potato Purple (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/23MzMhjR/Sweet-Potato-Purple.jpg", "stock": 0}, {"id": 336, "code": "VSPS", "name": "Sweet Potato Purple Skin (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/2YjCXhZW/Sweet-Potato-Purple-Skin.jpg", "stock": 0}, {"id": 337, "code": "VSPW", "name": "Sweet Potato White Flesh (Seasonal)", "cat": "Vegetables", "img": "https://i.ibb.co/HTkm8DSZ/Sweet-Potato-White-Flesh.jpg", "stock": 0}, {"id": 338, "code": "VSCB", "name": "Swiss Chard Bright Lights Mix", "cat": "Vegetables", "img": "https://i.ibb.co/HLBMfDxZ/Swiss-Chard-Bright-Lights-Mix.png", "stock": 100}, {"id": 339, "code": "VSCF", "name": "Swiss Chard Fordhook Giant", "cat": "Vegetables", "img": "https://i.ibb.co/6QK85qd/Swiss-Chard-Fordhook-Giant.png", "stock": 100}, {"id": 340, "code": "VTBL", "name": "Tomato Banana Legs", "cat": "Vegetables", "img": "https://i.ibb.co/xKZwd07k/Tomato-Banana-Legs.png", "stock": 100}, {"id": 341, "code": "VTBe", "name": "Tomato Beefsteak", "cat": "Vegetables", "img": "https://i.ibb.co/nsJqKrkr/Tomato-Beefsteak.png", "stock": 100}, {"id": 342, "code": "VTBZ", "name": "Tomato Black Zebra", "cat": "Vegetables", "img": "https://i.ibb.co/G4ftBfBN/Tomato-Black-Zebra.png", "stock": 100}, {"id": 343, "code": "VTCP", "name": "Tomato Cherokee Purple", "cat": "Vegetables", "img": "https://i.ibb.co/v6SftVt5/Tomato-Cherokee-Purple.png", "stock": 100}, {"id": 344, "code": "VTCB", "name": "Tomato Cherry Black", "cat": "Vegetables", "img": "https://i.ibb.co/21B7wH75/Tomato-Cherry-Black.png", "stock": 100}, {"id": 345, "code": "VTCM", "name": "Tomato Cherry Malawi Red", "cat": "Vegetables", "img": "https://i.ibb.co/zWpZtR1r/Tomato-Cherry-Malawi-Red.png", "stock": 100}, {"id": 346, "code": "VTCPI", "name": "Tomato Cherry Pink Ice", "cat": "Vegetables", "img": "https://i.ibb.co/gL2Mf6JK/Tomato-Cherry-Pink-Ice.png", "stock": 100}, {"id": 347, "code": "VTCY", "name": "Tomato Cherry Yellow", "cat": "Vegetables", "img": "https://i.ibb.co/PzzW30s0/Tomato-Cherry-Yellow.png", "stock": 100}, {"id": 348, "code": "VTCS", "name": "Tomato Chocolate Stripe", "cat": "Vegetables", "img": "https://i.ibb.co/j9mNJhbh/Tomato-Chocolate-Stripe.png", "stock": 100}, {"id": 349, "code": "VTCE", "name": "Tomato Cosmic Eclipse", "cat": "Vegetables", "img": "https://i.ibb.co/ns3np2Fp/Tomato-Cosmic-Eclipse.png", "stock": 100}, {"id": 350, "code": "VTGGo", "name": "Tomato Green Goddess", "cat": "Vegetables", "img": "https://i.ibb.co/YTQ8Dw8S/Tomato-Green-Goddess.png", "stock": 100}, {"id": 351, "code": "VTGS", "name": "Tomato Green Sausage", "cat": "Vegetables", "img": "https://i.ibb.co/ZzQyZtHy/Tomato-Green-Sausage.png", "stock": 100}, {"id": 352, "code": "VTHM", "name": "Tomato Heirloom Mix", "cat": "Vegetables", "img": "https://i.ibb.co/YBSj8xDh/Tomato-Heirloom-Mix.png", "stock": 100}, {"id": 353, "code": "VTMM", "name": "Tomato Money Maker", "cat": "Vegetables", "img": "https://i.ibb.co/hJrnjhb3/Tomato-Money-Maker.png", "stock": 100}, {"id": 354, "code": "VTOx", "name": "Tomato Oxheart", "cat": "Vegetables", "img": "https://i.ibb.co/L71kJ0x/Tomato-Oxheart.png", "stock": 100}, {"id": 355, "code": "VTPP", "name": "Tomato Purple Plum", "cat": "Vegetables", "img": "https://i.ibb.co/Y5QHv8J/Tomato-Purple-Plum.png", "stock": 100}, {"id": 356, "code": "VTRod", "name": "Tomato Rodade", "cat": "Vegetables", "img": "https://i.ibb.co/LXFd96XZ/Tomato-Rodade.png", "stock": 100}, {"id": 357, "code": "VTRom", "name": "Tomato Roma", "cat": "Vegetables", "img": "https://i.ibb.co/WpG3zFMb/Tomato-Roma.png", "stock": 100}, {"id": 358, "code": "VTGG", "name": "Turnip Green Globe", "cat": "Vegetables", "img": "https://i.ibb.co/MDdrj2cQ/Turnip-Green-Globe.png", "stock": 100}, {"id": 359, "code": "VTPT", "name": "Turnip Purple Top", "cat": "Vegetables", "img": "https://i.ibb.co/gMPJ3wmd/Turnip-Purple-Top.png", "stock": 100}, {"id": 360, "code": "VTSW", "name": "Turnip Snowball White", "cat": "Vegetables", "img": "https://i.ibb.co/KCg1h2x/Turnip-Snowball-White.png", "stock": 100}, {"id": 361, "code": "VTYG", "name": "Turnip Yellow Globe", "cat": "Vegetables", "img": "https://i.ibb.co/KpGbPSS8/Turnip-Yellow-Globe.png", "stock": 100}];
const LOGO = "https://i.ibb.co/Qj1wGkfJ/Trueleaf-Seeds-Logo.jpg";
const ADMIN_PASS = "TrueLeaf2026";

const PAL = {
  paper:"#f4ecd6", card:"#fffdf5", ink:"#2a2015", green:"#2d4a1e", greenMid:"#3d6b28",
  brown:"#5c3d1e", border:"#d8cfae", muted:"#8a7a5a", line:"#cabf98", terra:"#b5683c",
};
const CATS = ["All","Vegetables","Herbs","Flowers","Capsicums","Microgreens","Sprouts","Gourds","Fruit","Peanuts","Trees","Crop Cover & Lawns"];
const CATSTYLE = {
  "Vegetables":{icon:"🥕",c:"#3d6b28"}, "Herbs":{icon:"🌿",c:"#4f6b2e"}, "Flowers":{icon:"🌸",c:"#9c5786"},
  "Capsicums":{icon:"🌶️",c:"#9e3b2e"}, "Microgreens":{icon:"🌱",c:"#2f7d6b"}, "Sprouts":{icon:"🫘",c:"#3a5a8a"},
  "Gourds":{icon:"🎃",c:"#8a5a1e"}, "Fruit":{icon:"🍈",c:"#b5742a"}, "Peanuts":{icon:"🥜",c:"#8a6a3a"},
  "Trees":{icon:"🌳",c:"#2d4a1e"}, "Crop Cover & Lawns":{icon:"🌾",c:"#6b6a2e"}, "All":{icon:"📋",c:"#2d4a1e"},
};
const cstyle = (c) => CATSTYLE[c] || {icon:"🌿",c:"#4f6b2e"};
const byId = (id) => PRODUCTS.find(p=>p.id===id);

const SEED_ACCOUNTS = [
  {email:"demo@distributor.co.za", password:"demo", business:"Demo Distributors", contact:"Demo User", phone:"082 000 0000", addr1:"12 Market Street", suburb:"", city:"Johannesburg", province:"Gauteng", postal:"2001"},
  {email:"karoo@growers.co.za", password:"karoo", business:"Karoo Seedlings", contact:"Anika Botha", phone:"083 221 4567", addr1:"7 Church Square", suburb:"", city:"Graaff-Reinet", province:"Eastern Cape", postal:"6280"},
  {email:"highveld@seeds.co.za", password:"high", business:"Highveld Growers", contact:"Sipho Ndlovu", phone:"072 998 1122", addr1:"45 Lynnwood Road", suburb:"", city:"Pretoria", province:"Gauteng", postal:"0081"},
];

function seedOrders() {
  const d = (days) => { const t=new Date(); t.setDate(t.getDate()-days); return t.toISOString(); };
  const L = (arr) => arr.map(([id,qty])=>{const p=byId(id); return {id, code:p.code, name:p.name, cat:p.cat, qty};});
  const mk = (ref, acc, days, lines, status) => {
    const ls=L(lines); return {ref, when:d(days), email:acc.email, business:acc.business, contact:acc.contact,
      phone:acc.phone, addr:acc, lines:ls, totalUnits:ls.reduce((s,l)=>s+l.qty,0), lineCount:ls.length, status};
  };
  return [
    mk("TL-204881", SEED_ACCOUNTS[1], 1, [[0,40],[20,100],[55,60]], "Pending"),
    mk("TL-204862", SEED_ACCOUNTS[2], 2, [[5,50],[120,120],[200,80],[210,40]], "Pending"),
    mk("TL-204790", SEED_ACCOUNTS[1], 5, [[10,30],[14,30]], "Confirmed"),
    mk("TL-204710", SEED_ACCOUNTS[2], 9, [[3,100],[44,60],[88,40],[150,90],[260,50]], "Confirmed"),
    mk("TL-204502", SEED_ACCOUNTS[1], 18, [[7,70],[99,30]], "Confirmed"),
    mk("TL-204410", SEED_ACCOUNTS[2], 24, [[2,100],[33,100],[111,50]], "Confirmed"),
  ];
}

export default function App() {
  const [authReady, setAuthReady] = useState(false);
  const [user, setUser] = useState(null);
  const [adminAuthed, setAdminAuthed] = useState(false);
  const [view, setView] = useState("catalog");
  const [authMode, setAuthMode] = useState("login");
  const [cart, setCart] = useState({});            // keyed by id
  const [orders, setOrders] = useState([]);
  const [stock, setStock] = useState(() => { const s={}; PRODUCTS.forEach(p=>s[p.id]=p.stock); return s; });

  // ---- Firebase: session + profile ----
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, async (fb) => {
      if (!fb) { setUser(null); setAdminAuthed(false); setAuthReady(true); return; }
      try {
        const adminSnap = await getDoc(doc(db, "admins", fb.uid));
        if (adminSnap.exists()) { setAdminAuthed(true); setUser(null); setAuthReady(true); return; }
        const prof = await getDoc(doc(db, "distributors", fb.uid));
        setUser(prof.exists()
          ? { uid: fb.uid, ...prof.data() }
          : { uid: fb.uid, email: fb.email, business:"", contact:"", phone:"", addr1:"", suburb:"", city:"", province:"", postal:"" });
      } catch (e) { console.error(e); }
      setAuthReady(true);
    });
    return unsub;
  }, []);

  // ---- Firebase: live stock levels (after sign-in) ----
  useEffect(() => {
    if (!user && !adminAuthed) return;
    const unsub = onSnapshot(doc(db, "inventory", "levels"),
      (snap) => { if (snap.exists()) setStock(snap.data()); },
      () => {});
    return unsub;
  }, [user, adminAuthed]);
  const [search, setSearch] = useState("");
  const [cat, setCat] = useState("All");
  const [sort, setSort] = useState("default");
  const [light, setLight] = useState(null);
  const [toast, setToast] = useState("");

  const show = (m) => { setToast(m); setTimeout(()=>setToast(""), 1700); };

  const counts = useMemo(() => { const m={All:PRODUCTS.length}; PRODUCTS.forEach(p=>m[p.cat]=(m[p.cat]||0)+1); return m; }, []);
  const filtered = useMemo(() => {
    const q=search.trim().toLowerCase();
    let list = PRODUCTS.filter(p => (cat==="All"||p.cat===cat) && (q===""||p.name.toLowerCase().includes(q)||p.code.toLowerCase().includes(q)));
    if (sort==="az") list=[...list].sort((a,b)=>a.name.localeCompare(b.name));
    else if (sort==="za") list=[...list].sort((a,b)=>b.name.localeCompare(a.name));
    return list;
  }, [cat, search, sort]);

  const cartLines = useMemo(() => Object.keys(cart).filter(id=>cart[id]>0).map(id=>({...byId(+id), qty:cart[id]})).filter(x=>x.id!==undefined), [cart]);
  const totalUnits = cartLines.reduce((s,l)=>s+l.qty,0);
  const lineCount = cartLines.length;
  const setQty = (id, qty) => setCart(c=>({...c,[id]:Math.max(0,qty)}));
  const addOne = (id) => setCart(c=>({...c,[id]:(c[id]||0)+1}));

  const S = {
    page:{minHeight:"100vh", background:PAL.paper, color:PAL.ink, fontFamily:"ui-sans-serif,system-ui,sans-serif"},
    wrap:{maxWidth:1100, margin:"0 auto"},
    serif:{fontFamily:"Georgia,'Times New Roman',serif"},
    btn:{background:PAL.green, color:"#fff", border:"none", borderRadius:8, padding:"10px 16px", fontWeight:700, cursor:"pointer", fontSize:14},
    btnGhost:{background:"transparent", color:PAL.green, border:"1px solid "+PAL.green, borderRadius:8, padding:"9px 15px", fontWeight:700, cursor:"pointer", fontSize:14},
    input:{width:"100%", padding:"10px 12px", borderRadius:8, border:"1px solid "+PAL.border, background:"#fff", color:PAL.ink, fontSize:14, boxSizing:"border-box"},
    code:{fontFamily:"ui-monospace,Menlo,monospace", fontSize:11, fontWeight:700, letterSpacing:0.5, color:PAL.brown, background:"#f0e8cd", border:"1px solid "+PAL.line, borderRadius:4, padding:"2px 6px"},
  };

  // ---------- ADMIN ----------
  if (adminAuthed) return <Admin {...{orders,setOrders,stock,setStock,setAdminAuthed,S,PAL,show}} />;

  // ---------- AUTH ----------
  if (!authReady) return <Splash PAL={PAL}/>;
  if (!user) return <Auth {...{authMode,setAuthMode,S,PAL,show}} />;

  const submit = async () => {
    const base = {
      ref:"TL-"+Date.now().toString().slice(-6), when:new Date().toISOString(),
      uid:user.uid||null, email:user.email, business:user.business, contact:user.contact, phone:user.phone,
      addr:{addr1:user.addr1,suburb:user.suburb,city:user.city,province:user.province,postal:user.postal},
      lines:cartLines.map(l=>({id:l.id,code:l.code,name:l.name,cat:l.cat,qty:l.qty})),
      totalUnits, lineCount,
    };
    setView("thanks");
    try { await addDoc(collection(db,"orders"), {...base, createdAt:serverTimestamp(), status:"Pending"}); } catch(e){ console.error(e); }
    try { fetch("/.netlify/functions/submit-order", {method:"POST", headers:{"content-type":"application/json"}, body:JSON.stringify(base)}); } catch(e){ console.error(e); }
    setTimeout(()=>{ setCart({}); setView("catalog"); }, 2600);
  };

  const Nav = () => (
    <header style={{background:PAL.green, borderBottom:"3px solid "+PAL.brown}}>
      <div style={{...S.wrap, display:"flex", alignItems:"center", gap:16, padding:"10px 16px"}}>
        <div onClick={()=>setView("catalog")} style={{display:"flex", alignItems:"center", gap:10, cursor:"pointer"}}>
          <div style={{background:"#f4ecd6", borderRadius:7, padding:"3px 6px", display:"flex"}}>
            <img src={LOGO} alt="TrueLeaf Seed Co." style={{height:30, objectFit:"contain"}}/>
          </div>
          <div style={{color:"#f4ecd6"}}>
            <div style={{...S.serif, fontSize:16, fontWeight:700, lineHeight:1}}>Distributor Orders</div>
            <div style={{fontSize:10, letterSpacing:1.5, opacity:0.8, textTransform:"uppercase"}}>{user.business}</div>
          </div>
        </div>
        <div style={{flex:1}}/>
        <nav style={{display:"flex", gap:4, alignItems:"center"}}>
          <NB label="Catalogue" active={view==="catalog"} onClick={()=>setView("catalog")} PAL={PAL}/>
          <NB label={"My Order"+(lineCount?" ("+lineCount+")":"")} active={view==="order"} onClick={()=>setView("order")} PAL={PAL}/>
          <NB label="Profile" active={view==="profile"} onClick={()=>setView("profile")} PAL={PAL}/>
          <button onClick={()=>{signOut(auth);setCart({});setView("catalog");}} title="Log out" style={{background:"transparent", border:"none", color:"#e7d9a8", cursor:"pointer", padding:8, display:"flex"}}><LogOut size={18}/></button>
        </nav>
      </div>
    </header>
  );

  return (
    <div style={S.page}>
      {view==="thanks" ? (
        <Thanks S={S} PAL={PAL}/>
      ) : (
      <>
      <Nav/>
      <div style={{background:"#fbf4dd", borderBottom:"1px solid "+PAL.border, color:PAL.brown, fontSize:12, padding:"7px 16px", textAlign:"center"}}>
        <strong>Live ordering</strong> — your order is saved instantly to TrueLeaf and emailed to the team and to your own inbox.
      </div>

      {view==="catalog" && (
        <div style={{...S.wrap, padding:"16px"}}>
          <div style={{display:"flex", gap:8, flexWrap:"wrap", alignItems:"center", marginBottom:12}}>
            <div style={{position:"relative", flex:1, minWidth:200}}>
              <Search size={16} color={PAL.muted} style={{position:"absolute", left:11, top:12}}/>
              <input style={{...S.input, paddingLeft:34}} placeholder="Search by name or code…" value={search} onChange={e=>setSearch(e.target.value)}/>
            </div>
            <select style={{...S.input, width:"auto"}} value={sort} onChange={e=>setSort(e.target.value)}>
              <option value="default">Sort: Catalogue order</option><option value="az">Name: A–Z</option><option value="za">Name: Z–A</option>
            </select>
            <span style={{fontSize:12, color:PAL.muted, whiteSpace:"nowrap"}}>{filtered.length} of {PRODUCTS.length}</span>
          </div>
          <div style={{display:"flex", gap:6, flexWrap:"wrap", marginBottom:18}}>
            {CATS.map(c=>{ const active=cat===c; const cs=cstyle(c); return (
              <button key={c} onClick={()=>setCat(c)} style={{display:"inline-flex", alignItems:"center", gap:5, padding:"6px 12px", borderRadius:20, fontSize:12.5, fontWeight:700, cursor:"pointer", border:"1px solid "+(active?PAL.green:PAL.border), background:active?PAL.green:PAL.card, color:active?"#f4ecd6":PAL.ink}}>
                <span>{cs.icon}</span>{c}<span style={{opacity:0.6, fontWeight:600}}>{counts[c]||0}</span></button>
            );})}
          </div>
          <div style={{display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(168px,1fr))", gap:12}}>
            {filtered.map(p=>{ const cs=cstyle(p.cat); const qty=cart[p.id]||0; return (
              <div key={p.id} style={{background:PAL.card, border:"1px solid "+PAL.border, borderRadius:10, overflow:"hidden", display:"flex", flexDirection:"column"}}>
                <div onClick={()=>setLight(p)} style={{position:"relative", height:120, background:"#f0e8cd", cursor:"zoom-in", display:"flex", alignItems:"center", justifyContent:"center", overflow:"hidden"}}>
                  {p.img ? <img src={p.img} alt={p.name} style={{width:"100%", height:"100%", objectFit:"cover"}} onError={e=>{e.target.style.display="none";}}/> : <span style={{fontSize:38, opacity:0.5}}>{cs.icon}</span>}
                  <span style={{position:"absolute", top:6, right:6, background:"rgba(0,0,0,0.45)", color:"#fff", borderRadius:5, padding:3, display:"flex"}}><ZoomIn size={13}/></span>
                </div>
                <div style={{padding:"9px 10px 10px", flex:1, display:"flex", flexDirection:"column"}}>
                  <div style={{display:"flex", justifyContent:"space-between", alignItems:"flex-start", gap:6, marginBottom:6}}>
                    <span style={S.code}>{p.code}</span><span style={{fontSize:14}} title={p.cat}>{cs.icon}</span>
                  </div>
                  <p style={{...S.serif, margin:"0 0 10px", fontSize:13.5, fontWeight:700, lineHeight:1.25}}>{p.name}</p>
                  <div style={{marginTop:"auto"}}>
                    {qty>0
                      ? <Stepper qty={qty} onMinus={()=>setQty(p.id,qty-1)} onPlus={()=>addOne(p.id)} onSet={v=>setQty(p.id,v)} PAL={PAL}/>
                      : <button onClick={()=>addOne(p.id)} style={{...S.btn, width:"100%", padding:"8px", fontSize:13, display:"flex", alignItems:"center", justifyContent:"center", gap:6}}><Plus size={15}/> Add to order</button>}
                  </div>
                </div>
              </div>
            );})}
          </div>
          {filtered.length===0 && <div style={{textAlign:"center", padding:"3rem", color:PAL.muted}}>No products match that search.</div>}
        </div>
      )}

      {view==="order" && <OrderView {...{cartLines,totalUnits,setQty,addOne,setCart,setView,S,PAL,user,submit}} />}
      {view==="profile" && <Profile {...{user,setUser,S,PAL,show}} />}

      {light && (
        <div onClick={()=>setLight(null)} style={{position:"fixed", inset:0, background:"rgba(20,16,8,0.82)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:50, padding:20}}>
          <div onClick={e=>e.stopPropagation()} style={{background:PAL.card, borderRadius:12, maxWidth:420, width:"100%", overflow:"hidden", border:"3px solid "+PAL.brown}}>
            <div style={{position:"relative", background:"#f0e8cd", display:"flex", alignItems:"center", justifyContent:"center", minHeight:300}}>
              {light.img ? <img src={light.img} alt={light.name} style={{width:"100%", maxHeight:440, objectFit:"contain"}}/> : <div style={{textAlign:"center", padding:40}}><div style={{fontSize:72, opacity:0.5}}>{cstyle(light.cat).icon}</div><div style={{fontSize:12, color:PAL.muted, marginTop:8}}>Image coming soon</div></div>}
              <button onClick={()=>setLight(null)} style={{position:"absolute", top:10, right:10, background:"rgba(0,0,0,0.5)", color:"#fff", border:"none", borderRadius:"50%", width:32, height:32, cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"center"}}><X size={18}/></button>
            </div>
            <div style={{padding:16}}>
              <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:8}}>
                <span style={S.code}>{light.code}</span><span style={{fontSize:12, color:cstyle(light.cat).c, fontWeight:700}}>{cstyle(light.cat).icon} {light.cat}</span>
              </div>
              <h3 style={{...S.serif, margin:"0 0 14px", fontSize:20, color:PAL.green}}>{light.name}</h3>
              {(cart[light.id]||0)>0
                ? <Stepper qty={cart[light.id]} onMinus={()=>setQty(light.id,(cart[light.id]||0)-1)} onPlus={()=>addOne(light.id)} onSet={v=>setQty(light.id,v)} PAL={PAL} big/>
                : <button onClick={()=>{addOne(light.id); show("Added "+light.name);}} style={{...S.btn, width:"100%", padding:"11px", display:"flex", alignItems:"center", justifyContent:"center", gap:6}}><Plus size={16}/> Add to order</button>}
            </div>
          </div>
        </div>
      )}
      {toast && <div style={{position:"fixed", bottom:20, left:"50%", transform:"translateX(-50%)", background:PAL.ink, color:"#fff", padding:"10px 18px", borderRadius:8, fontSize:13, zIndex:60}}>{toast}</div>}
      </>
      )}
    </div>
  );
}

function authErr(e){ const c=(e&&e.code)||""; if(c.includes("invalid-cred")||c.includes("wrong-password")||c.includes("user-not-found")) return "No account matches that email and password."; if(c.includes("email-already-in-use")) return "An account with that email already exists."; if(c.includes("invalid-email")) return "That email address looks invalid."; if(c.includes("weak-password")) return "Password must be at least 6 characters."; if(c.includes("network")) return "Network error — check your connection."; return "Something went wrong. Please try again."; }

function Splash({PAL}) {
  return <div style={{minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:PAL.paper, color:PAL.muted, fontFamily:"ui-sans-serif,system-ui,sans-serif"}}>
    <div style={{textAlign:"center"}}><Leaf size={30} color={PAL.green}/><div style={{marginTop:10, fontSize:14}}>Loading…</div></div>
  </div>;
}

function NB({label,active,onClick,PAL}) {
  return <button onClick={onClick} style={{background:active?"#f4ecd6":"transparent", color:active?PAL.green:"#f4ecd6", border:"none", borderRadius:7, padding:"7px 12px", fontWeight:700, fontSize:13, cursor:"pointer"}}>{label}</button>;
}

function Stepper({qty,onMinus,onPlus,onSet,PAL,big}) {
  const sz=big?44:34;
  const bs={width:sz,height:sz,borderRadius:8,border:"1px solid "+PAL.border,background:"#fff",color:PAL.green,cursor:"pointer",fontWeight:700,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0};
  return (
    <div style={{display:"flex", alignItems:"center", gap:6, width:"100%"}}>
      <button onClick={onMinus} style={bs}><Minus size={big?18:15}/></button>
      <input value={qty} onChange={e=>onSet(parseInt(e.target.value.replace(/[^0-9]/g,''))||0)} style={{flex:1, textAlign:"center", height:sz, borderRadius:8, border:"1px solid "+PAL.border, background:"#fff", color:PAL.ink, fontWeight:700, fontSize:big?16:14, fontFamily:"ui-monospace,Menlo,monospace", width:"100%", boxSizing:"border-box"}}/>
      <button onClick={onPlus} style={bs}><Plus size={big?18:15}/></button>
    </div>
  );
}

function Field({label,v,on,S,PAL,type="text",span}) {
  return (
    <label style={{display:"block", gridColumn:span?"1 / -1":"auto"}}>
      <span style={{fontSize:11.5, fontWeight:700, color:PAL.brown, display:"block", marginBottom:4}}>{label}</span>
      <input type={type} value={v} onChange={e=>on(e.target.value)} style={S.input}/>
    </label>
  );
}

function Thanks({S,PAL}) {
  return (
    <div style={{minHeight:"100vh", background:PAL.green, display:"flex", alignItems:"center", justifyContent:"center", flexDirection:"column", color:"#f4ecd6", textAlign:"center", padding:20}}>
      <style>{"@keyframes pop{0%{transform:scale(0.4);opacity:0}55%{transform:scale(1.08)}100%{transform:scale(1);opacity:1}}@keyframes fade{from{opacity:0}to{opacity:1}}"}</style>
      <div style={{animation:"pop 0.6s ease-out", display:"inline-flex", width:96, height:96, borderRadius:"50%", background:"#f4ecd6", alignItems:"center", justifyContent:"center", marginBottom:22}}>
        <Sprout size={52} color={PAL.green}/>
      </div>
      <h1 style={{...S.serif, fontSize:34, margin:"0 0 8px", animation:"fade 0.8s ease-out"}}>Thank you!</h1>
      <p style={{fontSize:15, opacity:0.85, margin:0, maxWidth:340, animation:"fade 1.1s ease-out"}}>Your order has been sent to TrueLeaf. A confirmation is on its way to your email.</p>
    </div>
  );
}

function Auth({authMode,setAuthMode,S,PAL,show}) {
  const [f,setF] = useState({email:"",password:"",business:"",contact:"",phone:"",addr1:"",suburb:"",city:"",province:"",postal:""});
  const [err,setErr] = useState("");
  const up=(k,v)=>setF(s=>({...s,[k]:v}));
  const [busy,setBusy] = useState(false);
  const login=async()=>{ setErr(""); setBusy(true); try{ await signInWithEmailAndPassword(auth, f.email.trim(), f.password); }catch(e){ setErr(authErr(e)); } setBusy(false); };
  const register=async()=>{ setErr(""); for(const k of ["email","password","business","contact","phone","addr1","city","province","postal"]){ if(!String(f[k]).trim()){setErr("Please complete all required fields.");return;} } if(f.password.length<6){ setErr("Password must be at least 6 characters."); return; } setBusy(true); try{ const cred=await createUserWithEmailAndPassword(auth, f.email.trim(), f.password); await setDoc(doc(db,"distributors",cred.user.uid), {email:f.email.trim(),business:f.business,contact:f.contact,phone:f.phone,addr1:f.addr1,suburb:f.suburb,city:f.city,province:f.province,postal:f.postal,createdAt:serverTimestamp()}); }catch(e){ setErr(authErr(e)); setBusy(false); } };
  const demoLogin=async()=>{ setErr(""); setBusy(true); try{ await signInWithEmailAndPassword(auth,"demo@distributor.co.za","demo1234"); }catch(e){ try{ const c=await createUserWithEmailAndPassword(auth,"demo@distributor.co.za","demo1234"); await setDoc(doc(db,"distributors",c.user.uid),{email:"demo@distributor.co.za",business:"Demo Distributors",contact:"Demo User",phone:"082 000 0000",addr1:"12 Market Street",suburb:"",city:"Johannesburg",province:"Gauteng",postal:"2001",createdAt:serverTimestamp()}); }catch(e2){ setErr(authErr(e2)); setBusy(false);} } };

  return (
    <div style={{minHeight:"100vh", background:PAL.paper, color:PAL.ink, display:"flex", alignItems:"center", justifyContent:"center", padding:20, fontFamily:"ui-sans-serif,system-ui,sans-serif"}}>
      <div style={{width:"100%", maxWidth:authMode==="register"?560:400}}>
        <div style={{textAlign:"center", marginBottom:20}}>
          <img src={LOGO} alt="TrueLeaf Seed Co." style={{height:84, objectFit:"contain", marginBottom:6}}/>
          <p style={{margin:0, fontSize:12, letterSpacing:2, textTransform:"uppercase", color:PAL.muted}}>Distributor Ordering Portal</p>
        </div>
        <div style={{background:PAL.card, border:"1px solid "+PAL.border, borderRadius:12, padding:22}}>
          <div style={{display:"flex", gap:6, marginBottom:18, background:"#f0e8cd", borderRadius:9, padding:4}}>
            {["login","register"].map(m=>(<button key={m} onClick={()=>{setAuthMode(m);setErr("");}} style={{flex:1, padding:"8px", borderRadius:7, border:"none", cursor:"pointer", fontWeight:700, fontSize:13, background:authMode===m?PAL.green:"transparent", color:authMode===m?"#f4ecd6":PAL.muted}}>{m==="login"?"Log in":"Create account"}</button>))}
          </div>
          {authMode==="login" ? (
            <div style={{display:"grid", gap:10}}>
              <Field label="Email" v={f.email} on={v=>up("email",v)} S={S} PAL={PAL} type="email"/>
              <Field label="Password" v={f.password} on={v=>up("password",v)} S={S} PAL={PAL} type="password"/>
            </div>
          ) : (
            <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:10}}>
              <Field label="Business name *" v={f.business} on={v=>up("business",v)} S={S} PAL={PAL} span/>
              <Field label="Contact name *" v={f.contact} on={v=>up("contact",v)} S={S} PAL={PAL}/>
              <Field label="Phone *" v={f.phone} on={v=>up("phone",v)} S={S} PAL={PAL}/>
              <Field label="Email *" v={f.email} on={v=>up("email",v)} S={S} PAL={PAL} type="email"/>
              <Field label="Password *" v={f.password} on={v=>up("password",v)} S={S} PAL={PAL} type="password"/>
              <Field label="Delivery address *" v={f.addr1} on={v=>up("addr1",v)} S={S} PAL={PAL} span/>
              <Field label="Suburb" v={f.suburb} on={v=>up("suburb",v)} S={S} PAL={PAL}/>
              <Field label="City *" v={f.city} on={v=>up("city",v)} S={S} PAL={PAL}/>
              <Field label="Province *" v={f.province} on={v=>up("province",v)} S={S} PAL={PAL}/>
              <Field label="Postal code *" v={f.postal} on={v=>up("postal",v)} S={S} PAL={PAL}/>
            </div>
          )}
          {err && <p style={{color:"#b5341f", fontSize:13, margin:"12px 0 0"}}>{err}</p>}
          <button onClick={authMode==="login"?login:register} disabled={busy} style={{...S.btn, width:"100%", padding:12, marginTop:16, opacity:busy?0.7:1}}>{busy?"Please wait…":(authMode==="login"?"Log in":"Create account & continue")}</button>
          {authMode==="login" && <button onClick={demoLogin} disabled={busy} style={{...S.btnGhost, width:"100%", padding:10, marginTop:10}}>Explore as demo distributor</button>}
        </div>

        <div style={{textAlign:"center", marginTop:14}}>
          <span style={{color:PAL.muted, fontSize:12, display:"inline-flex", alignItems:"center", gap:5}}><ShieldCheck size={13}/> TrueLeaf admin (Greg): log in above with your admin email.</span>
        </div>
      </div>
    </div>
  );
}

function OrderView({cartLines,totalUnits,setQty,addOne,setCart,setView,S,PAL,user,submit}) {
  if (cartLines.length===0) return (
    <div style={{...S.wrap, padding:"60px 16px", textAlign:"center"}}>
      <ShoppingCart size={48} color={PAL.line} style={{margin:"0 auto 14px"}}/>
      <h2 style={{...S.serif, color:PAL.green, margin:"0 0 6px"}}>Your order is empty</h2>
      <p style={{color:PAL.muted, marginBottom:20}}>Add seed lines from the catalogue to build an order.</p>
      <button onClick={()=>setView("catalog")} style={S.btn}>Browse catalogue</button>
    </div>
  );
  return (
    <div style={{...S.wrap, padding:"16px"}}>
      <div style={{display:"flex", alignItems:"center", gap:10, marginBottom:14}}>
        <button onClick={()=>setView("catalog")} style={{...S.btnGhost, padding:"7px 11px", display:"flex", alignItems:"center", gap:5}}><ChevronLeft size={16}/> Add more</button>
        <h2 style={{...S.serif, color:PAL.green, margin:0, fontSize:22}}>My order</h2><div style={{flex:1}}/>
        <button onClick={()=>setCart({})} style={{background:"transparent", border:"none", color:PAL.muted, cursor:"pointer", fontSize:13, display:"flex", alignItems:"center", gap:5}}><Trash2 size={14}/> Clear</button>
      </div>
      <div style={{background:PAL.card, border:"1px solid "+PAL.border, borderRadius:10, overflow:"hidden"}}>
        <div style={{display:"flex", padding:"8px 14px", background:"#f0e8cd", fontSize:11, fontWeight:700, color:PAL.brown, textTransform:"uppercase", letterSpacing:0.5}}>
          <span style={{width:70}}>Code</span><span style={{flex:1}}>Seed</span><span style={{width:150, textAlign:"center"}}>Quantity</span><span style={{width:34}}/>
        </div>
        {cartLines.map(l=>(
          <div key={l.id} style={{display:"flex", alignItems:"center", padding:"10px 14px", borderTop:"1px solid "+PAL.border}}>
            <span style={{width:70, fontFamily:"ui-monospace,Menlo,monospace", fontSize:12, fontWeight:700, color:PAL.brown}}>{l.code}</span>
            <span style={{flex:1}}><span style={{...S.serif, fontWeight:700, fontSize:14}}>{l.name}</span><br/><span style={{fontSize:11, color:PAL.muted}}>{l.cat}</span></span>
            <div style={{width:150}}><Stepper qty={l.qty} onMinus={()=>setQty(l.id,l.qty-1)} onPlus={()=>addOne(l.id)} onSet={v=>setQty(l.id,v)} PAL={PAL}/></div>
            <button onClick={()=>setQty(l.id,0)} style={{width:34, background:"transparent", border:"none", color:PAL.muted, cursor:"pointer", display:"flex", justifyContent:"center"}}><X size={16}/></button>
          </div>
        ))}
        <div style={{display:"flex", justifyContent:"space-between", padding:"12px 14px", borderTop:"2px solid "+PAL.line, background:"#faf4de", fontWeight:700}}>
          <span>{cartLines.length} seed lines</span><span style={{fontFamily:"ui-monospace,Menlo,monospace"}}>{totalUnits} units total</span>
        </div>
      </div>
      <div style={{background:PAL.card, border:"1px solid "+PAL.border, borderRadius:10, padding:14, marginTop:14}}>
        <div style={{fontSize:11, fontWeight:700, color:PAL.brown, textTransform:"uppercase", letterSpacing:0.5, marginBottom:6}}>Deliver to</div>
        <div style={{fontSize:14, lineHeight:1.5}}><strong>{user.business}</strong> · {user.contact} · {user.phone}<br/>{user.addr1}{user.suburb?", "+user.suburb:""}, {user.city}, {user.province} {user.postal}</div>
        <button onClick={()=>setView("profile")} style={{...S.btnGhost, padding:"6px 11px", marginTop:10, fontSize:12}}>Edit delivery details</button>
      </div>
      <button onClick={submit} style={{...S.btn, width:"100%", padding:14, marginTop:16, fontSize:15, display:"flex", alignItems:"center", justifyContent:"center", gap:8}}><Send size={17}/> Submit order to TrueLeaf</button>
    </div>
  );
}

function Profile({user,setUser,S,PAL,show}) {
  const [f,setF]=useState({...user}); const up=(k,v)=>setF(s=>({...s,[k]:v}));
  const save=async()=>{ try{ await updateDoc(doc(db,"distributors",user.uid), {business:f.business,contact:f.contact,phone:f.phone,addr1:f.addr1,suburb:f.suburb,city:f.city,province:f.province,postal:f.postal}); setUser(f); show("Profile saved"); }catch(e){ show("Could not save profile"); } };
  return (
    <div style={{...S.wrap, padding:"16px", maxWidth:640}}>
      <h2 style={{...S.serif, color:PAL.green, margin:"0 0 14px", fontSize:22}}>Profile & delivery details</h2>
      <div style={{background:PAL.card, border:"1px solid "+PAL.border, borderRadius:10, padding:18}}>
        <div style={{display:"grid", gridTemplateColumns:"1fr 1fr", gap:12}}>
          <Field label="Business name" v={f.business} on={v=>up("business",v)} S={S} PAL={PAL} span/>
          <Field label="Contact name" v={f.contact} on={v=>up("contact",v)} S={S} PAL={PAL}/>
          <Field label="Phone" v={f.phone} on={v=>up("phone",v)} S={S} PAL={PAL}/>
          <Field label="Email" v={f.email} on={v=>up("email",v)} S={S} PAL={PAL}/>
          <Field label="Delivery address" v={f.addr1} on={v=>up("addr1",v)} S={S} PAL={PAL} span/>
          <Field label="Suburb" v={f.suburb} on={v=>up("suburb",v)} S={S} PAL={PAL}/>
          <Field label="City" v={f.city} on={v=>up("city",v)} S={S} PAL={PAL}/>
          <Field label="Province" v={f.province} on={v=>up("province",v)} S={S} PAL={PAL}/>
          <Field label="Postal code" v={f.postal} on={v=>up("postal",v)} S={S} PAL={PAL}/>
        </div>
        <button onClick={save} style={{...S.btn, marginTop:16}}>Save changes</button>
      </div>
    </div>
  );
}

// ===================== ADMIN =====================
function Admin({orders,setOrders,stock,setStock,setAdminAuthed,S,PAL,show}) {
  const [tab,setTab] = useState("overview");
  const [period,setPeriod] = useState("week");
  const [stockSearch,setStockSearch] = useState("");
  const [expanded,setExpanded] = useState(null);

  useEffect(() => {
    const qy = query(collection(db, "orders"), orderBy("createdAt", "desc"));
    return onSnapshot(qy, (snap) => setOrders(snap.docs.map(d => ({ id: d.id, ...d.data() }))));
  }, [setOrders]);
  useEffect(() => { (async () => {
    const lref = doc(db, "inventory", "levels");
    const snap = await getDoc(lref);
    if (!snap.exists()) { const lv={}; PRODUCTS.forEach(p=>lv[p.id]=p.stock); await setDoc(lref, lv); }
  })(); }, []);

  const daysAgo = (iso) => (Date.now()-new Date(iso).getTime())/86400000;
  const inPeriod = (o) => daysAgo(o.when) <= (period==="week"?7:31);
  const periodOrders = orders.filter(inPeriod);

  const byDist = useMemo(()=>{ const m={}; periodOrders.forEach(o=>{ if(!m[o.business]) m[o.business]={business:o.business, units:0, orders:0, cats:{}}; m[o.business].units+=o.totalUnits; m[o.business].orders+=1; o.lines.forEach(l=>{m[o.business].cats[l.cat]=(m[o.business].cats[l.cat]||0)+l.qty;}); }); return Object.values(m).sort((a,b)=>b.units-a.units); }, [periodOrders]);

  const totalUnits = periodOrders.reduce((s,o)=>s+o.totalUnits,0);
  const pendingCount = orders.filter(o=>o.status==="Pending").length;
  const outOfStock = Object.values(stock).filter(v=>v<=0).length;
  const lowStock = Object.values(stock).filter(v=>v>0&&v<=20).length;

  const confirmOrder = async (o) => {
    try {
      await runTransaction(db, async (tx) => {
        const oref = doc(db, "orders", o.id), lref = doc(db, "inventory", "levels");
        const os = await tx.get(oref);
        if (!os.exists() || os.data().status !== "Pending") return;
        const ls = await tx.get(lref); const lv = ls.exists() ? {...ls.data()} : {};
        o.lines.forEach(l => { lv[l.id] = Math.max(0, (lv[l.id] ?? 0) - l.qty); });
        tx.update(oref, { status: "Confirmed" });
        tx.set(lref, lv);
      });
      show("Order "+o.ref+" confirmed · stock updated");
    } catch (e) { show("Could not confirm order"); }
  };

  const downloadStock = () => {
    const rows = PRODUCTS.map(p=>({ID:p.id, Code:p.code, "Seed Name":p.name, Category:p.cat, Stock:stock[p.id]??0}));
    const ws = XLSX.utils.json_to_sheet(rows); const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Stock");
    const out = XLSX.write(wb, {type:"array", bookType:"xlsx"});
    const blob = new Blob([out], {type:"application/octet-stream"});
    const url = URL.createObjectURL(blob); const a=document.createElement("a");
    a.href=url; a.download="TrueLeaf_Stock_"+new Date().toISOString().slice(0,10)+".xlsx"; a.click(); URL.revokeObjectURL(url);
    show("Stock sheet downloaded");
  };
  const uploadStock = (e) => {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), {type:"array"});
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
        let n=0; setStock(prev=>{ const next={...prev}; rows.forEach(r=>{ if(r.ID!==undefined && r.Stock!==undefined){ next[+r.ID]=parseInt(r.Stock)||0; n++; } }); setDoc(doc(db,"inventory","levels"), next).catch(()=>{}); return next; });
        show(n+" stock levels updated from sheet");
      } catch(err){ show("Could not read that file"); }
    };
    reader.readAsArrayBuffer(file); e.target.value="";
  };

  const stockList = useMemo(()=>{ const q=stockSearch.trim().toLowerCase(); return PRODUCTS.filter(p=>q===""||p.name.toLowerCase().includes(q)||p.code.toLowerCase().includes(q)); }, [stockSearch]);

  const maxUnits = Math.max(1, ...byDist.map(d=>d.units));

  return (
    <div style={{minHeight:"100vh", background:PAL.paper, color:PAL.ink, fontFamily:"ui-sans-serif,system-ui,sans-serif"}}>
      <style>{"@keyframes grow{from{transform:scaleY(0)}to{transform:scaleY(1)}}@keyframes sway{0%,100%{transform:rotate(-2deg)}50%{transform:rotate(2deg)}}"}</style>
      <header style={{background:PAL.brown, borderBottom:"3px solid "+PAL.green}}>
        <div style={{maxWidth:1100, margin:"0 auto", display:"flex", alignItems:"center", gap:12, padding:"10px 16px"}}>
          <div style={{background:"#f4ecd6", borderRadius:7, padding:"3px 6px", display:"flex"}}><img src={LOGO} alt="TrueLeaf" style={{height:28, objectFit:"contain"}}/></div>
          <div style={{color:"#f4ecd6"}}><div style={{...S.serif, fontSize:16, fontWeight:700, lineHeight:1}}>Admin Dashboard</div><div style={{fontSize:10, letterSpacing:1.5, opacity:0.8, textTransform:"uppercase"}}>TrueLeaf Seed Co.</div></div>
          <div style={{flex:1}}/>
          <button onClick={()=>signOut(auth)} style={{background:"transparent", border:"none", color:"#f4ecd6", cursor:"pointer", display:"flex", alignItems:"center", gap:6, fontSize:13, fontWeight:700}}><LogOut size={16}/> Exit</button>
        </div>
      </header>

      <div style={{maxWidth:1100, margin:"0 auto", padding:"16px"}}>
        <div style={{display:"flex", gap:6, marginBottom:16}}>
          {[["overview","Overview",ClipboardList],["orders","Orders",Package],["stock","Stock",Sprout]].map(([k,lab,Ic])=>(
            <button key={k} onClick={()=>setTab(k)} style={{display:"flex", alignItems:"center", gap:6, padding:"9px 16px", borderRadius:8, border:"1px solid "+(tab===k?PAL.green:PAL.border), background:tab===k?PAL.green:PAL.card, color:tab===k?"#f4ecd6":PAL.ink, fontWeight:700, fontSize:14, cursor:"pointer"}}>
              <Ic size={16}/> {lab}{k==="orders"&&pendingCount>0?<span style={{background:"#b5341f", color:"#fff", borderRadius:10, fontSize:11, padding:"1px 7px"}}>{pendingCount}</span>:null}
            </button>
          ))}
        </div>

        {tab==="overview" && (
          <div>
            <div style={{display:"flex", alignItems:"center", gap:10, marginBottom:14, flexWrap:"wrap"}}>
              <h2 style={{...S.serif, color:PAL.green, margin:0, fontSize:22}}>The Order Garden</h2>
              <span style={{fontSize:12, color:PAL.muted}}>each distributor's plot grows with the units they've ordered</span>
              <div style={{flex:1}}/>
              <div style={{display:"flex", gap:4, background:"#f0e8cd", borderRadius:8, padding:4}}>
                {["week","month"].map(p=>(<button key={p} onClick={()=>setPeriod(p)} style={{padding:"6px 14px", borderRadius:6, border:"none", cursor:"pointer", fontWeight:700, fontSize:13, background:period===p?PAL.green:"transparent", color:period===p?"#f4ecd6":PAL.muted, textTransform:"capitalize"}}>{p==="week"?"This week":"This month"}</button>))}
              </div>
            </div>

            <div style={{display:"flex", gap:12, marginBottom:18, flexWrap:"wrap"}}>
              <Stat label="Orders" value={periodOrders.length} PAL={PAL} S={S}/>
              <Stat label="Units ordered" value={totalUnits} PAL={PAL} S={S}/>
              <Stat label="Awaiting confirmation" value={pendingCount} PAL={PAL} S={S} alert={pendingCount>0}/>
              <Stat label="Distributors active" value={byDist.length} PAL={PAL} S={S}/>
            </div>

            {/* THE GARDEN */}
            <div style={{background:"linear-gradient(180deg,#fbf6e4 0%,#f4ecd6 70%)", border:"1px solid "+PAL.border, borderRadius:12, padding:"24px 16px 0", overflow:"hidden"}}>
              {byDist.length===0 ? (
                <div style={{textAlign:"center", padding:"30px 0 50px", color:PAL.muted}}>No orders in this period yet — the garden is still dormant.</div>
              ) : (
                <div style={{display:"flex", alignItems:"flex-end", gap:18, justifyContent:"center", flexWrap:"wrap", minHeight:240}}>
                  {byDist.map((d,i)=>{
                    const h = 60 + Math.round((d.units/maxUnits)*150);
                    const topCat = Object.entries(d.cats).sort((a,b)=>b[1]-a[1])[0];
                    const col = topCat?cstyle(topCat[0]).c:PAL.greenMid;
                    return (
                      <div key={d.business} style={{display:"flex", flexDirection:"column", alignItems:"center", width:120}}>
                        <div style={{fontFamily:"ui-monospace,Menlo,monospace", fontWeight:700, fontSize:15, color:PAL.brown, marginBottom:4}}>{d.units}</div>
                        <Plant height={h} color={col} delay={i*0.12}/>
                        <div style={{width:"100%", height:14, background:"linear-gradient(180deg,#7a5230,#5c3d1e)", borderRadius:"3px 3px 0 0"}}/>
                        <div style={{textAlign:"center", marginTop:6, paddingBottom:18}}>
                          <div style={{...S.serif, fontSize:13, fontWeight:700, lineHeight:1.2}}>{d.business}</div>
                          <div style={{fontSize:11, color:PAL.muted}}>{d.orders} order{d.orders>1?"s":""} · {topCat?cstyle(topCat[0]).icon+" "+topCat[0]:""}</div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
            <p style={{fontSize:11.5, color:PAL.muted, marginTop:10, textAlign:"center"}}>Taller seedlings = more units ordered. Bloom colour reflects each distributor's most-ordered category.</p>
          </div>
        )}

        {tab==="orders" && (
          <div>
            <h2 style={{...S.serif, color:PAL.green, margin:"0 0 12px", fontSize:22}}>Orders in & out</h2>
            {orders.length===0 && <p style={{color:PAL.muted}}>No orders yet.</p>}
            <div style={{display:"grid", gap:10}}>
              {orders.map(o=>{
                const open = expanded===o.ref; const pend = o.status==="Pending";
                return (
                  <div key={o.ref} style={{background:PAL.card, border:"1px solid "+(pend?"#e0b94e":PAL.border), borderRadius:10, overflow:"hidden"}}>
                    <div onClick={()=>setExpanded(open?null:o.ref)} style={{display:"flex", alignItems:"center", gap:12, padding:"12px 14px", cursor:"pointer"}}>
                      <div style={{display:"flex", flexDirection:"column"}}>
                        <span style={{fontFamily:"ui-monospace,Menlo,monospace", fontWeight:700, color:PAL.brown, fontSize:13}}>{o.ref}</span>
                        <span style={{fontSize:11, color:PAL.muted}}>{new Date(o.when).toLocaleDateString("en-ZA",{day:"numeric",month:"short"})}</span>
                      </div>
                      <div style={{flex:1}}>
                        <div style={{...S.serif, fontWeight:700, fontSize:15}}>{o.business}</div>
                        <div style={{fontSize:12, color:PAL.muted}}>{o.lineCount} lines · {o.totalUnits} units</div>
                      </div>
                      <span style={{fontSize:12, fontWeight:700, padding:"4px 10px", borderRadius:20, background:pend?"#fbeec4":"#dcebd0", color:pend?"#86631a":PAL.green, display:"inline-flex", alignItems:"center", gap:5}}>
                        {pend?<AlertTriangle size={13}/>:<CheckCircle2 size={13}/>}{o.status}
                      </span>
                      {pend && <button onClick={(e)=>{e.stopPropagation(); confirmOrder(o);}} style={{...S.btn, padding:"7px 13px", fontSize:13}}>Confirm</button>}
                    </div>
                    {open && (
                      <div style={{borderTop:"1px solid "+PAL.border, padding:"12px 14px", background:"#fdfaf0"}}>
                        <div style={{fontSize:12.5, color:PAL.muted, marginBottom:10, lineHeight:1.5}}>
                          {o.contact} · {o.phone} · {o.email}<br/>
                          Deliver to: {o.addr.addr1}{o.addr.suburb?", "+o.addr.suburb:""}, {o.addr.city}, {o.addr.province} {o.addr.postal}<br/>
                          <span style={{color:PAL.brown}}>Email recipients: orders@trueleafseeds.co.za, wreford@donedigital.co.za, {o.email} (confirmation)</span>
                        </div>
                        <table style={{width:"100%", borderCollapse:"collapse", fontSize:13}}>
                          <thead><tr style={{background:"#f0e8cd", color:PAL.brown}}><th style={{textAlign:"left", padding:"5px 8px"}}>Code</th><th style={{textAlign:"left", padding:"5px 8px"}}>Seed</th><th style={{textAlign:"right", padding:"5px 8px"}}>Qty</th></tr></thead>
                          <tbody>{o.lines.map(l=>(<tr key={l.id} style={{borderBottom:"1px solid "+PAL.border}}><td style={{padding:"5px 8px", fontFamily:"ui-monospace,Menlo,monospace", fontWeight:700, color:PAL.brown}}>{l.code}</td><td style={{padding:"5px 8px"}}>{l.name}</td><td style={{padding:"5px 8px", textAlign:"right", fontFamily:"ui-monospace,Menlo,monospace", fontWeight:700}}>{l.qty}</td></tr>))}</tbody>
                        </table>
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </div>
        )}

        {tab==="stock" && (
          <div>
            <div style={{display:"flex", alignItems:"center", gap:10, marginBottom:14, flexWrap:"wrap"}}>
              <h2 style={{...S.serif, color:PAL.green, margin:0, fontSize:22}}>Stock</h2><div style={{flex:1}}/>
              <button onClick={downloadStock} style={{...S.btnGhost, display:"flex", alignItems:"center", gap:6, padding:"8px 13px"}}><Download size={15}/> Download sheet</button>
              <label style={{...S.btn, display:"inline-flex", alignItems:"center", gap:6, padding:"9px 13px", cursor:"pointer"}}>
                <Upload size={15}/> Upload sheet
                <input type="file" accept=".xlsx,.xls" onChange={uploadStock} style={{display:"none"}}/>
              </label>
            </div>
            <div style={{display:"flex", gap:12, marginBottom:14, flexWrap:"wrap"}}>
              <Stat label="Products" value={PRODUCTS.length} PAL={PAL} S={S}/>
              <Stat label="Low stock (≤20)" value={lowStock} PAL={PAL} S={S} alert={lowStock>0}/>
              <Stat label="Out of stock" value={outOfStock} PAL={PAL} S={S} alert={outOfStock>0}/>
            </div>
            <div style={{position:"relative", marginBottom:12}}>
              <Search size={16} color={PAL.muted} style={{position:"absolute", left:11, top:12}}/>
              <input style={{...S.input, paddingLeft:34}} placeholder="Find a seed by name or code…" value={stockSearch} onChange={e=>setStockSearch(e.target.value)}/>
            </div>
            <div style={{background:PAL.card, border:"1px solid "+PAL.border, borderRadius:10, overflow:"hidden"}}>
              <div style={{display:"flex", padding:"8px 14px", background:"#f0e8cd", fontSize:11, fontWeight:700, color:PAL.brown, textTransform:"uppercase", letterSpacing:0.5}}>
                <span style={{width:70}}>Code</span><span style={{flex:1}}>Seed</span><span style={{width:120, textAlign:"center"}}>In stock</span>
              </div>
              <div style={{maxHeight:480, overflowY:"auto"}}>
                {stockList.slice(0,200).map(p=>{ const s=stock[p.id]??0; const warn=s<=0?"#b5341f":s<=20?"#b5742a":PAL.green; return (
                  <div key={p.id} style={{display:"flex", alignItems:"center", padding:"8px 14px", borderTop:"1px solid "+PAL.border}}>
                    <span style={{width:70, fontFamily:"ui-monospace,Menlo,monospace", fontSize:12, fontWeight:700, color:PAL.brown}}>{p.code}</span>
                    <span style={{flex:1}}><span style={{...S.serif, fontWeight:700, fontSize:13.5}}>{p.name}</span> <span style={{fontSize:11, color:PAL.muted}}>· {p.cat}</span></span>
                    <div style={{width:120, display:"flex", justifyContent:"center"}}>
                      <input value={s} onChange={e=>{const v=parseInt(e.target.value.replace(/[^0-9]/g,''))||0; setStock(st=>({...st,[p.id]:v}));}} onBlur={()=>{ setDoc(doc(db,"inventory","levels"), stock).catch(()=>{}); }} style={{width:70, textAlign:"center", padding:"6px", borderRadius:7, border:"1px solid "+PAL.border, fontFamily:"ui-monospace,Menlo,monospace", fontWeight:700, color:warn}}/>
                    </div>
                  </div>
                );})}
                {stockList.length>200 && <div style={{padding:"10px 14px", fontSize:12, color:PAL.muted, textAlign:"center"}}>Showing first 200 — search to narrow down.</div>}
              </div>
            </div>
            <p style={{fontSize:11.5, color:PAL.muted, marginTop:10}}>Edit numbers inline, or download the sheet, update it in Excel, and upload it back. Keep the <strong>ID</strong> column intact — it's how the upload matches rows (the codes aren't all unique).</p>
          </div>
        )}
      </div>
    </div>
  );
}

function Stat({label,value,PAL,S,alert}) {
  return (
    <div style={{background:alert?"#fbeec4":PAL.card, border:"1px solid "+(alert?"#e0b94e":PAL.border), borderRadius:10, padding:"12px 18px", minWidth:130, flex:"1 1 130px"}}>
      <div style={{fontSize:26, fontWeight:800, fontFamily:"Georgia,serif", color:alert?"#86631a":PAL.green}}>{value}</div>
      <div style={{fontSize:11.5, color:PAL.muted, fontWeight:600, textTransform:"uppercase", letterSpacing:0.5}}>{label}</div>
    </div>
  );
}

function Plant({height,color,delay}) {
  const w=70; const stemX=w/2; const top=10;
  return (
    <svg width={w} height={height} viewBox={"0 0 "+w+" "+height} style={{display:"block", transformOrigin:"bottom center", animation:"grow 0.9s ease-out "+delay+"s both"}}>
      <path d={"M"+stemX+" "+height+" C "+(stemX-6)+" "+(height*0.6)+", "+(stemX+5)+" "+(height*0.4)+", "+stemX+" "+top} stroke="#3d6b28" strokeWidth="3.5" fill="none" strokeLinecap="round"/>
      {[0.72,0.52,0.34].map((ly,i)=>{ const y=top+(height-top)*ly; const dir=i%2===0?1:-1; return (
        <g key={i} style={{transformOrigin:stemX+"px "+y+"px", animation:"sway 4s ease-in-out "+(i*0.3)+"s infinite"}}>
          <path d={"M"+stemX+" "+y+" Q "+(stemX+dir*26)+" "+(y-14)+", "+(stemX+dir*30)+" "+(y+4)+" Q "+(stemX+dir*16)+" "+(y+8)+", "+stemX+" "+y+" Z"} fill="#4f8a36"/>
        </g>
      );})}
      <circle cx={stemX} cy={top} r="9" fill={color}/>
      <circle cx={stemX} cy={top} r="4" fill="#fff8e6" opacity="0.85"/>
    </svg>
  );
}
