
# Cube Add-in

## Workflow

### 1. Bronbestanden aanpassen
- `customUI14.xml` – Ribbon-interface aanpassen.
- `Cube.bas` – VBA-functionaliteit aanpassen.

### 2. Open de projectmap

```bash
cd ~/Documents/Coding/cube-addin
```

### 3. Ribbon injecteren

```bash
bash inject_ribbon.sh Cube-addin.xlam
```

### 4. Resultaat controleren

Bij succes verschijnt:

```text
Klaar: lint toegevoegd aan Cube-addin.xlam
(back-up: Cube-addin.xlam.bak)
```

Het script maakt automatisch een back-up van de bestaande add-in.

### 5. Excel herstarten

- Sluit Excel volledig af.
- Open Excel opnieuw.
- Laad de Cube-add-in.
- Controleer of de Ribbon correct wordt weergegeven.

---

## Projectstructuur

```text
cube-addin/
│
├── Cube-addin.xlam
├── Cube.bas
├── customUI14.xml
├── inject_ribbon.sh
└── README.md
```

## Belangrijk

- Iedere wijziging aan `customUI14.xml` vereist een nieuwe injectie.
- VBA-wijzigingen in `Cube.bas` moeten afzonderlijk in het VBA-project van de add-in worden verwerkt.
- Het injectiescript past alleen de Ribbon XML aan, tenzij het expliciet anders is ingericht.
- De vorige versie wordt automatisch bewaard als `.bak`.

## Terminal command

```bash
bash inject_ribbon.sh Cube-addin.xlam
```
