from contextlib import contextmanager
from pathlib import Path
from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree

import matplotlib.pyplot as plt
import pandas as pd
from openpyxl import load_workbook
from matplotlib.ticker import FuncFormatter


FOLDER = Path(__file__).parent
OUTPUT_FOLDER = FOLDER / "eredmenyek"
CORE_PROPERTIES_NS = "{http://purl.org/dc/terms/}"

SPORTAGAK = [
	("Amerikai futball", ("FOOTBALL",)),
	("Baseball", ("BASEBALL",)),
	("Softball", ("SOFTBALL",)),
	("Kosárlabda", ("BASKETBALL",)),
	("Labdarúgás", ("SOCCER",)),
	("Röplabda", ("VOLLEYBALL",)),
	("Hoki", ("HOCKEY",)),
	("Asztalitenisz", ("TABLE TENNIS",)),
	("Tenisz", ("TENNIS",)),
	("Golf", ("GOLF",)),
	("Gimnasztika", ("GYMNASTICS",)),
	("Birkózás", ("WRESTLING",)),
	("Lacrosse", ("LACROSSE",)),
	("Rögbi", ("RUGBY",)),
	("Cheerleading", ("CHEERLEADING",)),
	("Úszás", ("SWIMMING",)),
	("Kerékpározás", ("BICYCLE", "CYCLING")),
	("Gördeszkázás", ("SKATEBOARD",)),
	("Síelés", ("SKIING", "SKI EQUIPMENT")),
	("Snowboard", ("SNOWBOARD",)),
]


def egysegesites(value):
	if pd.isna(value):
		return ""
	try:
		return str(int(float(value)))
	except (TypeError, ValueError):
		return str(value).strip()


@contextmanager
def neiisadatokmegnyit(file_path):
	javitott_adatok = None
	try:
		adatok = load_workbook(file_path, read_only=True, data_only=True)
	except ValueError as error:
		if "could not read properties" not in str(error):
			raise

		javitott_adatok = BytesIO()
		with ZipFile(file_path) as ahelye, ZipFile(
			javitott_adatok, "w", compression=ZIP_DEFLATED
		) as javitottremelem:
			for cucc in ahelye.infolist():
				acucc = ahelye.read(cucc.filename)
				if cucc.filename == "docProps/core.xml":
					root = ElementTree.fromstring(acucc)
					for name in ("created", "modified"):
						datuma = root.find(CORE_PROPERTIES_NS + name)
						if datuma is not None:
							root.remove(datuma)
					acucc = ElementTree.tostring(
						root, encoding="utf-8", xml_declaration=True
					)
				javitottremelem.writestr(cucc, acucc)
		javitott_adatok.seek(0)
		try:
			adatok = load_workbook(javitott_adatok, read_only=True, data_only=True)
		except Exception:
			javitott_adatok.close()
			raise

	try:
		yield adatok
	finally:
		adatok.close()
		if javitott_adatok is not None:
			javitott_adatok.close()


def beolvasas(file_path):
	with neiisadatokmegnyit(file_path) as workbook:
		sheet = workbook["NEISS_FMT"]
		sheet.reset_dimensions()
		rows = sheet.iter_rows(values_only=True)
		next(rows, None)
		labels = {}
		for row in rows:
			if row[0] == "PROD" and row[1] is not None and row[3] is not None:
				labels[egysegesites(row[1])] = str(row[3]).upper()
		return labels


def sportrend(product_label): 
	if not isinstance(product_label, str):
		return None

	for sport, keywords in SPORTAGAK:
		if any(keyword in product_label for keyword in keywords):
			return sport

	if "SPORTS AND RECREATIONAL ACTIVITY" in product_label:
		return "Egyéb sport/rekreáció"
	return None


def serulesbetolt(files):
	osszadat = []
	required_columns = ["CPSC_Case_Number", "Sex", "Product_1", "Weight"]
	for file_path in files:
		print(f"Feldolgozás: {file_path.stem.replace('neiss', '')}...", flush=True)
		product_labels = beolvasas(file_path)
		with neiisadatokmegnyit(file_path) as workbook:
			sheet = workbook.worksheets[0]
			sheet.reset_dimensions()
			rows = sheet.iter_rows(values_only=True)
			headers = next(rows)
			oszlop_indexes = {name: headers.index(name) for name in required_columns}
			esetek = []
			for row in rows:
				if row[oszlop_indexes["CPSC_Case_Number"]] is None:
					continue
				product_kod = egysegesites(row[oszlop_indexes["Product_1"]])
				sport = sportrend(product_labels.get(product_kod))
				weight = row[oszlop_indexes["Weight"]]
				if sport is None or weight is None:
					continue
				try:
					weight = float(weight)
				except (TypeError, ValueError):
					continue
				nem = row[oszlop_indexes["Sex"]]
				nem = {0: "Ismeretlen", 1: "Férfi", 2: "Nő"}.get(nem, "Ismeretlen")
				esetek.append({"sport": sport, "nem": nem, "Weight": weight})
			osszadat.extend(esetek)

	if not osszadat:
		raise ValueError("Nem találtam feldolgozható NEISS-adatot.")
	return pd.DataFrame(osszadat, columns=["sport", "nem", "Weight"])


def diagment(data):
	nemenkent = data[data["nem"].isin(["Férfi", "Nő"])].pivot_table(
		index="sport", columns="nem", values="Weight", aggfunc="sum", fill_value=0
	)
	nemenkent = nemenkent.reindex(columns=["Férfi", "Nő"], fill_value=0)
	nemenkent = nemenkent.loc[nemenkent.sum(axis=1) > 0]
	nemenkent = nemenkent.loc[nemenkent.sum(axis=1).sort_values().index]
	fig, ax = plt.subplots(figsize=(13, max(7, len(nemenkent) * 0.62)))
	fig.patch.set_facecolor("white")
	ax.set_facecolor("#f7f9fa")
	nemenkent.plot(kind="barh", ax=ax, color={"Férfi": "#177e89", "Nő": "#e07a5f"})
	for bars in ax.containers:
		labels = [
			f"{bar.get_width():,.1f}".replace(",", " ").replace(".", ",")
			for bar in bars
		]
		ax.bar_label(bars, labels=labels, padding=4, fontsize=8)
	ax.set_xlim(0, nemenkent.to_numpy().max() * 1.24)
	ax.set_title("Sportserülések sportág és nem szerint (2016–2025)", fontsize=16, pad=14)
	ax.set_xlabel("Sérülések száma (millió)")
	ax.set_ylabel("")
	ax.xaxis.set_major_formatter(FuncFormatter(lambda value, _: f"{value / 1_000_000:g}"))
	ax.grid(axis="x", color="#dce3e6", linewidth=0.8)
	ax.set_axisbelow(True)
	ax.tick_params(axis="y", labelsize=9)
	ax.legend(title="Nem", frameon=False, loc="lower right")
	fig.tight_layout()
	fig.savefig(OUTPUT_FOLDER / "sportag_es_nem_szerint.png", dpi=160)
	plt.close(fig)

def main():
	files = sorted(FOLDER.glob("neiss20*.xlsx"))
	if not files:
		raise FileNotFoundError(
			"Nem találtam neiss20xx.xlsx nevű fájlt. Tedd az Excel-fájlokat a valmai.py mellé."
		)

	OUTPUT_FOLDER.mkdir(exist_ok=True)
	serulesek = serulesbetolt(files)
	if serulesek.empty:
		raise ValueError(
			"Nem sikerült sportághoz sorolható esetet találni. Ellenőrizd az Excel-fájlok termékkódjait."
		)

	nemeenkent = (
		serulesek.groupby(["sport", "nem"], as_index=False)["Weight"]
		.sum()
		.rename(columns={"sport": "Sportág", "nem": "Nem", "Weight": "Becsült esetszám"})
	)
	nemeenkent.to_csv(OUTPUT_FOLDER / "sportag_es_nem.csv", index=False, encoding="utf-8-sig")
	diagment(serulesek)

	print(f"Besorolt NEISS-esetek: {len(serulesek):,}".replace(",", " "))
	print(f"Országos becslés a Weight mező alapján: {serulesek['Weight'].sum():,.0f}".replace(",", " "))
	print(f"Elkészült eredmények mappája: {OUTPUT_FOLDER}")

if __name__ == "__main__":
	main()
