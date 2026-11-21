"""
src/ui.py - Reusable Tailwind CSS UI Components for Streamlit
Utilizes Tailwind CDN injected inside st.components.v1.html iframes.
"""

import streamlit as st
import streamlit.components.v1 as components

TAILWIND_CDN = '<script src="https://cdn.tailwindcss.com"></script>'
GOOGLE_FONTS = '<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">'

def render_header(title: str, subtitle: str, badge_text: str = "MCA DA Capstone"):
    """
    Renders an executive, professional header hero banner styled with Tailwind CSS.
    """
    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        {GOOGLE_FONTS}
        <style>
            body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 0; background-color: transparent; }}
        </style>
    </head>
    <body>
        <div class="bg-gradient-to-r from-blue-900 via-indigo-900 to-slate-900 text-white p-6 rounded-2xl shadow-lg border border-blue-800/40">
            <div class="flex items-center justify-between">
                <div>
                    <span class="inline-block px-3 py-1 bg-blue-500/20 text-blue-300 text-xs font-semibold rounded-full uppercase tracking-wider mb-2 border border-blue-400/30">
                        {badge_text}
                    </span>
                    <h1 class="text-2xl font-bold tracking-tight text-white mb-1">
                        {title}
                    </h1>
                    <p class="text-sm text-slate-300 font-normal">
                        {subtitle}
                    </p>
                </div>
                <div class="hidden sm:flex flex-col items-end text-xs text-slate-400 border-l border-slate-700/60 pl-6">
                    <span class="font-semibold text-slate-200">Gourav Kumar (2500320140055)</span>
                    <span class="text-blue-300 font-medium">Project ID: P083</span>
                    <span>Mentor: Mr. Dinesh Sereno</span>
                    <span class="text-slate-400 text-[10px]">Batch: ABES_OND26_MCA 2nd Year_Data Analytics_Batch 01</span>
                </div>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=140, scrolling=False)


def render_kpi_cards(kpi_data: list):
    """
    Renders a responsive flexbox grid of high-contrast Tailwind KPI metric cards.
    kpi_data format:
    [
        {"title": "Blended CAC", "value": "$71.40", "sub": "-12.5% vs Q3", "status": "positive"},
        {"title": "Total Spend", "value": "$248,500", "sub": "24 Months Total", "status": "neutral"},
        {"title": "Customers Acquired", "value": "3,480", "sub": "+18.2% YoY", "status": "positive"},
        {"title": "Avg Payback Period", "value": "8.4 Mo", "sub": "Safe (<12 Mo)", "status": "positive"}
    ]
    """
    card_items = ""
    for item in kpi_data:
        # Determine status pill color
        status = item.get("status", "neutral")
        if status == "positive":
            pill_class = "text-emerald-700 bg-emerald-50 border-emerald-200"
            dot_color = "bg-emerald-500"
        elif status == "negative" or status == "critical":
            pill_class = "text-rose-700 bg-rose-50 border-rose-200"
            dot_color = "bg-rose-500"
        else:
            pill_class = "text-blue-700 bg-blue-50 border-blue-200"
            dot_color = "bg-blue-500"

        card_items += f"""
        <div class="flex-1 min-w-[200px] bg-white p-5 rounded-xl border border-slate-200 shadow-sm hover:shadow-md transition duration-150">
            <div class="flex items-center justify-between mb-2">
                <span class="text-xs font-semibold text-slate-500 uppercase tracking-wide">{item['title']}</span>
                <span class="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium border {pill_class}">
                    <span class="w-1.5 h-1.5 mr-1.5 rounded-full {dot_color}"></span>
                    {item.get('sub', '')}
                </span>
            </div>
            <div class="text-2xl font-bold text-slate-900 tracking-tight">{item['value']}</div>
            {f"<div class='text-xs text-slate-400 mt-1'>{item['detail']}</div>" if 'detail' in item else ""}
        </div>
        """

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        {GOOGLE_FONTS}
        <style>
            body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 0; background-color: transparent; }}
        </style>
    </head>
    <body>
        <div class="flex flex-wrap gap-4 py-2">
            {card_items}
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=115, scrolling=False)


def render_insight_box(box_type: str, title: str, text: str):
    """
    Renders an informative or alert insight box.
    box_type: 'warning', 'success', 'info'
    """
    if box_type == "warning":
        container_class = "bg-amber-50 border-amber-300 text-amber-900"
        title_class = "text-amber-800"
        icon = "⚠️"
    elif box_type == "success":
        container_class = "bg-emerald-50 border-emerald-300 text-emerald-900"
        title_class = "text-emerald-800"
        icon = "✅"
    elif box_type == "critical":
        container_class = "bg-rose-50 border-rose-300 text-rose-900"
        title_class = "text-rose-800"
        icon = "🚨"
    else:
        container_class = "bg-blue-50 border-blue-300 text-blue-900"
        title_class = "text-blue-800"
        icon = "ℹ️"

    html_code = f"""
    <!DOCTYPE html>
    <html>
    <head>
        {TAILWIND_CDN}
        {GOOGLE_FONTS}
        <style>
            body {{ font-family: 'Inter', sans-serif; margin: 0; padding: 0; background-color: transparent; }}
        </style>
    </head>
    <body>
        <div class="p-4 rounded-xl border {container_class} flex items-start gap-3 my-2 shadow-xs">
            <span class="text-xl leading-none">{icon}</span>
            <div class="text-sm">
                <span class="font-bold {title_class} block mb-0.5">{title}</span>
                <span class="opacity-90">{text}</span>
            </div>
        </div>
    </body>
    </html>
    """
    components.html(html_code, height=85, scrolling=False)
