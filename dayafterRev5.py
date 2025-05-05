import pandas as pd
import os
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from dataclasses import dataclass
from matplotlib import rcParams
rcParams['font.family'] = 'sans-serif'
rcParams['font.sans-serif'] = ['Segoe UI', 'DejaVu Sans', 'Arial']

@dataclass
class SummaryMetrics:
    vol_4g: float
    vol_5g: float
    total_volume: float
    offload: float
    peak_hour: str
    peak_4g: int
    peak_5g: int
    peak_total: int
    tput_metrics: dict
    top_cells: dict

class DataProcessor:
    @staticmethod
    def load_excel_files(folder_path):
        """Load all Excel files from a folder into a DataFrame"""
        if not os.path.exists(folder_path):
            raise FileNotFoundError(f"Folder not found: {folder_path}")
        
        files = [os.path.join(folder_path, f) 
                 for f in os.listdir(folder_path) if f.endswith('.xlsx')]
        if not files:
            raise ValueError(f"No Excel files found in {folder_path}")
            
        return pd.concat([pd.read_excel(file) for file in files], ignore_index=True)

    @staticmethod
    def process_tech_data(folder_path, config):
        """Process technology-specific data based on configuration"""
        df = DataProcessor.load_excel_files(folder_path)
        df = df.rename(columns=config['columns_rename'])
        df = df.drop(columns=config['columns_to_drop'])
        #df.columns = config['column_names']
        # if 'column_order' in config:
        #     df = df[config['column_order']]
        df['Tech'] = config['tech']
        df['Grupo'] = df['Grupo'].str.upper()
        df['Date'] = pd.to_datetime(df['Date'])  # Convert to datetime
        
        # Apply conversions
        conversions = {
            'VolumeGB': 1e6,
            'TputDLMB': 1e3,
            'TputULMB': 1e3
        }
        for col, divisor in conversions.items():
            if col in df.columns:
                df[col] = df[col] / divisor
                
        return df

class ReportGenerator:
    def __init__(self, df, output_path, groups=None):
        self.df = df.dropna(subset=['Users', 'Disp', 'TputDLMB', 'TputULMB'])
        self.output_path = output_path
        self.groups = groups or self.df['Grupo'].dropna().unique()
        self.metrics = ['VolumeGB', 'TputDLMB', 'TputULMB', 'Users', 'Disp', 'acc']
        self.aggregations = {
            'VolumeGB': 'sum',
            'Users': 'sum',
            'TputDLMB': 'mean',
            'TputULMB': 'mean',
            'Disp': 'mean',
            'acc' : 'mean'
        }

    def _create_site_metric_plots(self, group, tech, pdf):
        """Generate time series plots per site for each metric"""
        df_group_tech = self.df[(self.df['Grupo'] == group) & (self.df['Tech'] == tech)]
        
        if df_group_tech.empty:
            return

        for metric in self.metrics:
            fig, ax = plt.subplots(figsize=(12, 6))
            fig.suptitle(f'{metric} for Grupo {group} - {tech}', fontsize=14)
            
            # Aggregate data by Site and Date
            df_agg = df_group_tech.groupby(['Date', 'Site']).agg(
                {metric: self.aggregations[metric]}
            ).reset_index()

            # Plot each site's time series
            for site in df_agg['Site'].unique():
                site_data = df_agg[df_agg['Site'] == site]
                ax.plot(site_data['Date'], site_data[metric], 
                        label=site, marker='o', linestyle='-', markersize=4)

            ax.set_xlabel('Date', fontsize=10)
            ax.set_ylabel(metric, fontsize=10)
            ax.tick_params(axis='x', rotation=45, labelsize=8)
            ax.tick_params(axis='y', labelsize=8)
            ax.legend(title='Site', bbox_to_anchor=(1.05, 1), 
                     loc='upper left', fontsize=8)
            ax.grid(True, alpha=0.3)
            
            plt.tight_layout()
            pdf.savefig(fig)
            plt.close()
    


    def _calculate_summary_metrics(self, df_group):
        """Calculate all summary metrics"""
        # Volume calculations
        vol_4g = df_group[df_group['Tech'] == '4G']['VolumeGB'].sum()
        vol_5g = df_group[df_group['Tech'] == '5G']['VolumeGB'].sum()
        total_volume = vol_4g + vol_5g
        offload = vol_5g / total_volume if total_volume > 0 else 0

        # Initialize default values
        metrics = SummaryMetrics(
            vol_4g=vol_4g,
            vol_5g=vol_5g,
            total_volume=total_volume,
            offload=offload,
            peak_hour="N/A",
            peak_4g=0,
            peak_5g=0,
            peak_total=0,
            tput_metrics={},
            top_cells={}
        )

        try:
            # Peak user calculations
            user_pivot = df_group.pivot_table(
                index='Date',
                columns='Tech',
                values='Users',
                aggfunc='sum',
                fill_value=0
            )
            user_pivot['Total'] = user_pivot['4G'] + user_pivot['5G']
            
            if not user_pivot.empty:
                peak_time = user_pivot['Total'].idxmax()
                df_peak = df_group[df_group['Date'] == peak_time]

                metrics.peak_hour = f"{peak_time.hour:02d}h"
                metrics.peak_4g = int(user_pivot.loc[peak_time, '4G'])
                metrics.peak_5g = int(user_pivot.loc[peak_time, '5G'])
                metrics.peak_total = int(user_pivot.loc[peak_time, 'Total'])

                # Throughput calculations
                metrics.tput_metrics = {
                    '4G DL': df_peak[df_peak['Tech'] == '4G']['TputDLMB'].mean(),
                    '4G UL': df_peak[df_peak['Tech'] == '4G']['TputULMB'].mean(),
                    '5G DL': df_peak[df_peak['Tech'] == '5G']['TputDLMB'].mean(),
                    '5G UL': df_peak[df_peak['Tech'] == '5G']['TputULMB'].mean(),
                }

                # Top cells calculation
                metrics.top_cells = (df_peak.groupby('Cell')['Users']
                                    .sum()
                                    .nlargest(5)
                                    .to_dict())

        except Exception as e:
            print(f"Error calculating metrics: {str(e)}")
        
        return metrics

    def _create_summary_page(self, group, pdf):
        """Create compact single-page summary and worst cells page"""
        df_group = self.df[self.df['Grupo'] == group]
        metrics = self._calculate_summary_metrics(df_group)

        # Create summary page (existing code)
        fig = plt.figure(figsize=(10, 10), layout="constrained")
        ax = fig.add_subplot()
        ax.axis('off')

        styles = {
            'header': {'fontsize': 12, 'fontweight': 'bold', 'color': '#2e74b5'},
            'subheader': {'fontsize': 11, 'fontweight': 'semibold', 'color': '#404040'},
            'value': {'fontsize': 11, 'color': '#404040'},
            'highlight': {'fontsize': 11, 'fontweight': 'bold', 'color': '#c00000'},
            'spacer': {'fontsize': 1}
        }

        text_content = [
            (f'Relatório: {group}', 'header'),
            ('', 'spacer'),
            ('Volume de Dados (GB):', 'subheader'),
            (f"4G: {metrics.vol_4g:.2f} | 5G: {metrics.vol_5g:.2f}", 'value'),
            (f"Total: {metrics.total_volume:.2f}", 'highlight'),
            ('', 'spacer'),
            ('Offload 5G:', 'subheader'),
            (f"{metrics.offload:.2%}", 'highlight'),
            ('', 'spacer'),
            (f"Pico de Usuários @ {metrics.peak_hour}", 'subheader'),
            (f"4G: {metrics.peak_4g:,} | 5G: {metrics.peak_5g:,}", 'value'),
            (f"Total: {metrics.peak_total:,}", 'highlight'),
        ]

        if metrics.tput_metrics:
            text_content += [
                ('', 'spacer'),
                ('Throughput Médio (Mbps):', 'subheader'),
                (f"4G DL: {metrics.tput_metrics.get('4G DL', 0):.1f} UL: {metrics.tput_metrics.get('4G UL', 0):.1f}", 'value'),
                (f"5G DL: {metrics.tput_metrics.get('5G DL', 0):.1f} UL: {metrics.tput_metrics.get('5G UL', 0):.1f}", 'value'),
            ]

        if metrics.top_cells:
            text_content += [
                ('', 'spacer'),
                ('Top Células (usuários):', 'subheader')
            ]
            for cell, users in metrics.top_cells.items():
                truncated_cell = cell[:15] + '...' if len(cell) > 15 else cell
                text_content.append((f"{truncated_cell}: {users:,}", 'value'))

        y_start = 0.93
        line_height = 0.038
        current_y = y_start

        for content, style_type in text_content:
            if style_type == 'spacer':
                current_y -= line_height
                continue
            ax.text(0.5, current_y, content,
                    ha='center', va='center',
                    transform=fig.transFigure,
                    **styles[style_type])
            current_y -= line_height
            if current_y < 0.05:
                break

        pdf.savefig(fig)
        plt.close()

        # New page for worst cells by accessibility
        fig = plt.figure(figsize=(10, 10), layout="constrained")
        ax = fig.add_subplot()
        ax.axis('off')

        # Get 5 worst cells by 'acc'
        if 'acc' in df_group.columns:
            worst_cells = df_group.nsmallest(5, 'acc')[['Cell', 'acc']].values.tolist()
        else:
            worst_cells = []

        text_content = [
            (f'5 Piores Células por Acessibilidade ({group})', 'header'),
            ('', 'spacer'),
        ]

        if worst_cells:
            text_content += [
                ('Célula          Acessibilidade (%)', 'subheader'),
                ('', 'spacer'),
            ]
            for cell, acc in worst_cells:
                truncated_cell = cell[:5] + '...' if len(cell) > 15 else cell
                text_content.append((f"{truncated_cell}: {acc:.2f}", 'value'))
        else:
            text_content.append(('Dados de acessibilidade não disponíveis.', 'value'))

        current_y = y_start
        for content, style_type in text_content:
            if style_type == 'spacer':
                current_y -= line_height
                continue
            ax.text(0.5, current_y, content,
                    ha='center', va='center',
                    transform=fig.transFigure,
                    **styles[style_type])
            current_y -= line_height
            if current_y < 0.05:
                break

        pdf.savefig(fig)
        plt.close()
        
    def _create_time_series_plots(self, group, pdf):
        """Generate time series comparison plots"""
        df_group = self.df[self.df['Grupo'] == group]
        fig, axes = plt.subplots(2, 3, figsize=(20, 10))
        fig.suptitle(f'Metrics for Grupo {group}', fontsize=16)
        fig.subplots_adjust(hspace=0.5, wspace=0.3)

        for idx, metric in enumerate(self.metrics):
            ax = axes[idx//3, idx%3]
            self._plot_dual_axis(ax, df_group, metric)
            
        plt.tight_layout()
        pdf.savefig(fig)
        plt.close()

    def _plot_dual_axis(self, ax, df_group, metric):
        """Plot dual axis time series for a metric"""
        # 4G data
        df_4g = df_group[df_group['Tech'] == '4G'].groupby('Date').agg(
            {metric: self.aggregations[metric]}).reset_index()
        ax.plot(df_4g['Date'], df_4g[metric], color='b', label='4G')
        ax.set_ylabel(f'4G {metric}', color='b')
        ax.tick_params(axis='y', labelcolor='b')
        
        # 5G data
        ax2 = ax.twinx()
        df_5g = df_group[df_group['Tech'] == '5G'].groupby('Date').agg(
            {metric: self.aggregations[metric]}).reset_index()
        ax2.plot(df_5g['Date'], df_5g[metric], color='r', label='5G')
        ax2.set_ylabel(f'5G {metric}', color='r')
        ax.tick_params(axis='x', rotation=90)
        ax2.tick_params(axis='y', labelcolor='r')
        
        ax.set_title(metric)
        ax.grid(True)

    def _create_boxplots(self, group, tech, pdf):
        """Generate boxplots for different metrics"""
        df_tech = self.df[(self.df['Grupo'] == group) & (self.df['Tech'] == tech)]
        #df_tech = df_tech.sort_values(by=metric, ascending=False)
        for metric in ['TputDLMB', 'TputULMB', 'Users']:
            fig = self._single_boxplot(df_tech, metric, tech, group)
            pdf.savefig(fig)
            plt.close()

    def _single_boxplot(self, df, metric, tech, group):
        """Create single boxplot figure"""
        fig, ax = plt.subplots(figsize=(20, 5))
        sorted_cells = df.groupby('Cell')[metric].max().sort_values().index
        
        ax.boxplot([df[df['Cell'] == cell][metric] 
                   for cell in sorted_cells], vert=True)
        ax.set_xticklabels(sorted_cells, rotation=90, fontsize=8)
        ax.set_xlabel('Cell')
        ax.set_ylabel(metric)
        ax.set_title(f'Boxplot for {metric} - {tech} - {group}')
        plt.tight_layout()
        return fig

    def _create_cell_users_facet_plots(self, group, tech, pdf):
        """Generate facet plots of Users vs Date for each cell, up to 24 per page in two columns"""
        df_group_tech = self.df[(self.df['Grupo'] == group) & (self.df['Tech'] == tech)]
        if df_group_tech.empty:
            return

        cells = df_group_tech['Cell'].dropna().unique()
        if len(cells) == 0:
            return

        # Page layout configuration
        rows_per_page = 24
        cols_per_page = 3
        plots_per_page = rows_per_page * cols_per_page

        # Generate pages with facet plots
        for page_num, i in enumerate(range(0, len(cells), plots_per_page)):
            current_cells = cells[i:i + plots_per_page]
            
            fig, axes = plt.subplots(rows_per_page, cols_per_page, figsize=(15, 30))
            fig.suptitle(f'Users per Cell - {group} - {tech} (Page {page_num + 1})', fontsize=14)
            axes_flat = axes.flatten()

            for j, cell in enumerate(current_cells):
                ax = axes_flat[j]
                cell_data = df_group_tech[df_group_tech['Cell'] == cell]
                df_agg = cell_data.groupby('Date')['Users'].sum().reset_index()
                
                ax.plot(df_agg['Date'], df_agg['Users'], 
                        marker='o', markersize=2, linestyle='-', linewidth=1)
                ax.set_title(cell[:15] + ('...' if len(cell) > 15 else ''), fontsize=6)
                ax.tick_params(axis='x', labelsize=4, rotation=45)
                ax.tick_params(axis='y', labelsize=4)
                ax.grid(True, alpha=0.3)

            # Disable empty subplots
            for j in range(len(current_cells), plots_per_page):
                axes_flat[j].axis('off')

            plt.tight_layout(rect=[0, 0, 1, 0.97])  # Adjust layout to accommodate suptitle
            pdf.savefig(fig)
            plt.close()

    def generate_report(self):
        """Main method to generate full report"""
        with PdfPages(self.output_path) as pdf:
            for group in self.groups:
                self._create_summary_page(group, pdf)
                self._create_time_series_plots(group, pdf)
                
                for tech in ['4G', '5G']:
                    self._create_boxplots(group, tech, pdf)
                    self._create_site_metric_plots(group, tech, pdf)  # New addition
                    self._create_cell_users_facet_plots(group, tech, pdf)  # New method call

# Configuration for data processing
tech_configs = {
    '4G': {
        'columns_to_drop': ['Detentora', 'Vendor'],
        'column_names': ['Date', 'Grupo', 'Site', 'Cell',  'TputDLMB', 'TputULMB', 'Disp', 
                         'VolumeGB', 'PRB_DL', 'Users', 'acc'],
        'column_order': ['Date', 'Grupo', 'Site', 'Cell', 'TputDLMB',
                        'TputULMB', 'Disp', 'VolumeGB', 'Users','acc'],
        'columns_rename': {'TIM_THROU_USER_PDCP_DL (Kbps)': 'TputDLMB',
                           'TIM_THROU_USER_PDCP_UL (Kbps)': 'TputULMB',
                           'TIM_DISP_COUNTER_TOTAL (%)': 'Disp',
                           'TIM_VOLUME_TOTAL_DLUL_ALLOP (KB)' : 'VolumeGB',
                           'TIM_PRB_UTIL_MEAN_DL (%)': 'PRB_DL',
                            'TIM_USERS_RRC_CONN_MAX_SUM (Units)': 'Users',
                            'TIM_ACC (%)': 'acc',
                            'eNodeB': 'Site'},
        'tech': '4G'
    },
    '5G': {
        'columns_to_drop': ['Fornecedor', 'gNodeB Name'],
        'column_names': ['Date', 'Grupo', 'Site', 'Cell', 'acc', 'Disp', 'Users',
                         'VolumeGB','TputULMB','TputDLMB'],
        'column_order': ['Date', 'Grupo', 'Site', 'Cell', 'TputDLMB',
                        'TputULMB', 'Disp', 'VolumeGB', 'Users','acc'],
        'columns_rename':{'TIM_ACC (%)' : 'acc',
                        'TIM_DISP_COUNTER_TOTAL (%)': 'Disp',
                        'TIM_USERS_RRC_CONN_MAX_SUM (Units)': 'Users',
                        'TIM_VOLUME_TOTAL_DLUL_ALLOP (KB)' : 'VolumeGB',
                        'TIM_THROU_USER_UL (Kbps)'	: 'TputULMB',
                        'TIM_THROU_USER_DL (Kbps)' : 'TputDLMB',
                        'gNodeB': 'Site'},                                              
        'tech': '5G'
    }
}

# Main execution
if __name__ == "__main__":
    # Load and process data
    df_4g = DataProcessor.process_tech_data('4G', tech_configs['4G'])
    df_5g = DataProcessor.process_tech_data('5G', tech_configs['5G'])

    # Combine and clean data
    df_4g.to_excel('dados_2204.xlsx', index=False)
    combined_df = pd.concat([df_4g, df_5g], ignore_index=True)
    combined_df = combined_df.drop_duplicates()
    #combined_df = combined_df[(combined_df['Date'] >= '2025-03-22 12:00:00') & (combined_df['Date'] <= '2025-03-22 21:00:00')]
    # Generate report
    report = ReportGenerator(
        combined_df,
        output_path='report_GARANHUNS_MAI25.pdf',
        groups=list(combined_df['Grupo'].unique())
    )
    report.generate_report()

    df_4g['Cell'] == '4G-CE05CW-26-1A'