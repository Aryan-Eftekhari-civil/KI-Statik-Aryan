import numpy as np
import matplotlib.pyplot as plt


def solve_heat_conduction():
    """
    Löst das stationäre Wärmeleitungsproblem für eine Wand
    """
    # Gegebene Parameter
    d = 0.25  # Wanddicke in m (25 cm)
    k = 0.15  # Wärmeleitfähigkeit in W/(m·K)
    Ti = 22  # Innentemperatur in °C
    Ta = 2  # Außentemperatur in °C
    hi = 1 / 0.13  # Wärmeübergangskoeffizient innen in W/(m²·K)
    ha = 1 / 0.04  # Wärmeübergangskoeffizient außen in W/(m²·K)

    print("Wärmeleitungsproblem - Stationäre Wärmeströmung")
    print("=" * 50)
    print(f"Wanddicke: {d * 1000:.0f} mm")
    print(f"Wärmeleitfähigkeit: {k:.2f} W/(m·K)")
    print(f"Innentemperatur: {Ti:.0f} °C")
    print(f"Außentemperatur: {Ta:.0f} °C")
    print(f"Wärmeübergangskoeff. innen: {hi:.2f} W/(m²·K)")
    print(f"Wärmeübergangskoeff. außen: {ha:.2f} W/(m²·K)")
    print("-" * 50)

    # Lösung der Differentialgleichung T(x) = C1*x + C2
    # Randbedingungen:
    # Bei x=0: -k*dT/dx = hi*(Ti - T(0))
    # Bei x=d: -k*dT/dx = ha*(T(d) - Ta)

    # Aufstellen des Gleichungssystems
    # T(x) = C1*x + C2
    # dT/dx = C1

    # Randbedingung bei x=0: -k*C1 = hi*(Ti - C2)
    # Randbedingung bei x=d: -k*C1 = ha*(C1*d + C2 - Ta)

    # Aus der ersten Gleichung: C2 = Ti + k*C1/hi
    # Einsetzen in die zweite Gleichung:
    # -k*C1 = ha*(C1*d + Ti + k*C1/hi - Ta)

    # Auflösen nach C1:
    C1 = ha* hi * (Ta - Ti) / (k * ha + k * hi + d * ha * hi)
    C2 = (k * ha* Ta + k * hi * Ti + d * ha * hi * Ti) / (k * ha + k * hi + d * ha * hi)

    print(f"Konstanten der linearen Funktion T(x) = C1*x + C2:")
    print(f"C1 = {C1:.2f} K/m")
    print(f"C2 = {C2:.2f} K")

    # Temperaturprofil berechnen
    x = np.linspace(0, d, 100)
    T = C1 * x + C2

    # Oberflächentemperaturen
    T_surface_inner = C2  # T(0)
    T_surface_outer = C1 * d + C2  # T(d)

    print(f"\nOberflächentemperaturen:")
    print(f"Innere Oberfläche (x=0): {T_surface_inner:.2f} °C")
    print(f"Äußere Oberfläche (x={d:.2f}m): {T_surface_outer:.2f} °C")

    # Wärmestromdichte berechnen
    q = -k * C1  # q = -k * dT/dx

    print(f"\nWärmestromdichte:")
    print(f"q = {q:.2f} W/m²")

    # Gesamtwärmewiderstand
    R_total = 1 / hi + d / k + 1 / ha
    q_alternative = (Ti - Ta) / R_total

    print(f"\nVerifikation über Gesamtwärmewiderstand:")
    print(f"R_gesamt = 1/hi + d/k + 1/ha = {R_total:.4f} m²K/W")
    print(f"q = (Ti - Ta) / R_gesamt = {q_alternative:.2f} W/m²")

    # Visualisierung
    plt.figure(figsize=(12, 8))

    # Temperaturprofil
    plt.subplot(2, 2, 1)
    plt.plot(x * 1000, T, 'b-', linewidth=2, label='Temperaturprofil')
    plt.axhline(y=Ti, color='r', linestyle='--', alpha=0.7, label=f'Innentemperatur ({Ti}°C)')
    plt.axhline(y=Ta, color='c', linestyle='--', alpha=0.7, label=f'Außentemperatur ({Ta}°C)')
    plt.scatter([0, d * 1000], [T_surface_inner, T_surface_outer],
                c=['red', 'blue'], s=100, zorder=5)
    plt.xlabel('Position x [mm]')
    plt.ylabel('Temperatur T [°C]')
    plt.title('Temperaturverlauf in der Wand')
    plt.grid(True, alpha=0.3)
    plt.legend()

    # Wärmestromdichte (konstant)
    plt.subplot(2, 2, 2)
    plt.axhline(y=q, color='g', linewidth=3, label=f'q = {q:.2f} W/m²')
    plt.xlabel('Position x [mm]')
    plt.ylabel('Wärmestromdichte q [W/m²]')
    plt.title('Wärmestromdichte (konstant)')
    plt.grid(True, alpha=0.3)
    plt.legend()
    plt.xlim(0, d * 1000)

    # Temperaturgradient
    plt.subplot(2, 2, 3)
    gradient = np.full_like(x, C1)
    plt.plot(x * 1000, gradient, 'r-', linewidth=2, label=f'dT/dx = {C1:.2f} K/m')
    plt.xlabel('Position x [mm]')
    plt.ylabel('Temperaturgradient dT/dx [K/m]')
    plt.title('Temperaturgradient (konstant)')
    plt.grid(True, alpha=0.3)
    plt.legend()

    # Wärmewiderstände
    plt.subplot(2, 2, 4)
    resistances = [1 / hi, d / k, 1 / ha]
    labels = ['R_i = 1/h_i', 'R_wall = d/k', 'R_a = 1/h_a']
    colors = ['lightblue', 'orange', 'lightgreen']

    bars = plt.bar(labels, resistances, color=colors, alpha=0.7)
    plt.ylabel('Wärmewiderstand [m²K/W]')
    plt.title('Wärmewiderstände')
    plt.xticks(rotation=45)

    # Werte auf den Balken anzeigen
    for bar, resistance in zip(bars, resistances):
        plt.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.01,
                 f'{resistance:.3f}', ha='center', va='bottom')

    plt.tight_layout()
    plt.show()

    # Zusätzliche Berechnungen
    print(f"\n" + "=" * 50)
    print("ZUSÄTZLICHE ERGEBNISSE:")
    print("=" * 50)

    # Temperaturdifferenzen
    delta_T_inner = Ti - T_surface_inner
    delta_T_wall = T_surface_inner - T_surface_outer
    delta_T_outer = T_surface_outer - Ta

    print(f"Temperaturdifferenzen:")
    print(f"ΔT_innen = {delta_T_inner:.2f} K")
    print(f"ΔT_Wand = {delta_T_wall:.2f} K")
    print(f"ΔT_außen = {delta_T_outer:.2f} K")
    print(f"ΔT_gesamt = {Ti - Ta:.2f} K")

    # Einzelwiderstände in Prozent
    R_inner_percent = (1 / hi) / R_total * 100
    R_wall_percent = (d / k) / R_total * 100
    R_outer_percent = (1 / ha) / R_total * 100

    print(f"\nAnteil der Wärmewiderstände:")
    print(f"Innerer Übergang: {R_inner_percent:.1f}%")
    print(f"Wand: {R_wall_percent:.1f}%")
    print(f"Äußerer Übergang: {R_outer_percent:.1f}%")

    return {
        'C1': C1, 'C2': C2,
        'T_surface_inner': T_surface_inner,
        'T_surface_outer': T_surface_outer,
        'q': q,
        'R_total': R_total,
        'x': x, 'T': T
    }


# Hauptprogramm ausführen
if __name__ == "__main__":
    results = solve_heat_conduction()

    print(f"\n" + "=" * 50)
    print("ANTWORT AUF DIE FRAGE:")
    print("=" * 50)
    print(f"Das Temperaturprofil zum Zeitpunkt t = ∞ (stationärer Zustand)")
    print(f"stellt sich als lineare Funktion dar:")
    print(f"T(x) = {results['C1']:.2f} * x + {results['C2']:.2f}")
    print(f"wobei x in Metern und T in °C angegeben ist.")
    print(f"\nBei einer Wanddicke von 25 cm:")
    print(f"- Innere Oberflächentemperatur: {results['T_surface_inner']:.2f} °C")
    print(f"- Äußere Oberflächentemperatur: {results['T_surface_outer']:.2f} °C")