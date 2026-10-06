"""Shared deterministic colours for native charts and rendered report figures."""

from colorsys import hsv_to_rgb

SERIES_COLOURS = ("4F81BD", "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646",
                  "235789", "A23B72", "007F5F", "B5651D", "6262A8", "555555")


def series_colour(index: int) -> str:
    if index < len(SERIES_COLOURS):
        return SERIES_COLOURS[index]
    # Extend without cycling back to a colour already used by another series.
    used = set(SERIES_COLOURS)
    for candidate in range(index - len(SERIES_COLOURS) + 1):
        hue = ((candidate + 1) * 0.618033988749895) % 1.0
        rgb = hsv_to_rgb(hue, 0.65, 0.72)
        colour = "".join(f"{round(component * 255):02X}" for component in rgb)
        while colour in used:
            colour = f"{(int(colour, 16) + 1) % 0x1000000:06X}"
        used.add(colour)
    return colour
