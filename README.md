Project Title
Astro-Logistics: Automated Site Mobilization & Heavy Lifting Tracker for Lunar Habitats

Summary
Building the Artemis lunar base isn’t just an aerospace challenge; it’s an unprecedented heavy-industrial construction nightmare. Our solution, Astro-Logistics, bridges the gap between aerospace engineering and on-ground heavy construction management. We developed an automated Python-based tool that extracts critical data from isometric piping drawings and equipment manifests, automatically calculates Lunar-gravity weight adjustments, and generates dynamic Project Management dashboards. This tool optimizes site mobilization, sequences heavy lifting operations, and tracks piping spool installations, ensuring safe and efficient construction in a micro-gravity environment.

What exactly does it do?
Astro-Logistics acts as the digital "Site Manager" for lunar construction.

Data Extraction: It reads raw structural and mechanical engineering data (simulated as CSV input from NASA blueprints).

Gravity-Adjusted Engineering Logic: It automatically converts Earth-weights to Lunar-weights, identifying rigging requirements for massive equipment (e.g., 170-ton nitrogen storage tanks).

Automated Tracking Generation: Using pandas and openpyxl, it generates a color-coded, multi-sheet Excel dashboard in seconds. This workbook includes:

Project Dashboard: Overall mobilization progress and launch payload limits.

Heavy Lifting Schedule: Prioritized sequence for moving critical infrastructure.

Piping & Spools Map: Tracking the installation of critical life-support and cooling pipelines.

Safety Alerts: It conditionally formats and flags any equipment that exceeds safe lunar lifting capacities.

How we addressed the challenge
Most space exploration concepts focus on getting to the Moon. We focused on what happens next. Drawing on real-world experience in managing heavy industrial mega-projects (like chemical complexes and heavy lifting operations), we applied terrestrial construction logic to space exploration. We realized that manual tracking of piping and heavy lifts on the Moon would lead to critical supply chain bottlenecks. Our code automates the translation of complex blueprints into actionable mobilization schedules, directly addressing the logistical bottleneck of establishing a permanent off-world presence.
