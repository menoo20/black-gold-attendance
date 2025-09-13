# GitHub Pages Deployment Guide for Black Gold Attendance Dashboard

## Step 1: Create GitHub Repository

1. Go to [github.com](https://github.com) and log in
2. Click "New Repository" (green button)
3. Repository settings:
   - **Name**: `attendance-analysis` or `black-gold-attendance`
   - **Description**: "Black Gold Training Attendance Analysis Dashboard"
   - **Visibility**: Choose Public (for GitHub Pages) or Private (if company prefers)
   - **Initialize**: Don't check any boxes (we have existing code)
4. Click "Create Repository"

## Step 2: Push Your Code to GitHub

After creating the repository, run these commands in your terminal:

```bash
# Add GitHub as remote origin
git remote add origin https://github.com/YOUR-USERNAME/attendance-analysis.git

# Switch to main branch (GitHub Pages default)
git branch -M main

# Push all your code
git push -u origin main
```

**Replace `YOUR-USERNAME` with your actual GitHub username!**

## Step 3: Enable GitHub Pages

1. Go to your repository on GitHub
2. Click "Settings" tab (top right of repository)
3. Scroll down to "Pages" in the left sidebar
4. Under "Source", select "Deploy from a branch"
5. Choose "main" branch and "/ (root)" folder
6. Click "Save"

## Step 4: Configure for Dashboard Access

GitHub Pages will serve your site at: `https://YOUR-USERNAME.github.io/attendance-analysis/`

The master dashboard will be at: `https://YOUR-USERNAME.github.io/attendance-analysis/master_dashboard.html`

## Step 5: Update README with Live URL

After deployment, update the README_COMPANY.md with your actual URL:

1. Replace `https://your-username.github.io/attendance-analysis/` 
2. With your actual URL: `https://YOUR-USERNAME.github.io/attendance-analysis/master_dashboard.html`

## Step 6: Set Custom Domain (Optional)

If Black Gold has a custom domain:

1. In repository Settings > Pages
2. Add custom domain: `attendance.blackgold.com` (example)
3. Create CNAME file in repository root with domain name
4. Configure DNS with company IT team

## Commands to Run Now:

```bash
# 1. First, switch to main branch for GitHub Pages
git branch -M main

# 2. Add your GitHub repository as remote
git remote add origin https://github.com/YOUR-USERNAME/REPO-NAME.git

# 3. Push everything to GitHub
git push -u origin main

# 4. Then follow steps 3-6 above in GitHub web interface
```

## Expected File Structure on GitHub:

```
Repository Root:
├── master_dashboard.html          # Main dashboard (entry point)
├── README_COMPANY.md              # Company documentation
├── weeks/                         # All week data
│   ├── week_31Aug-4Sep/
│   │   └── dashboard_week_31Aug-4Sep.html
│   └── week_7Sep-11Sep/
│       └── dashboard_week_7Sep-11Sep.html
├── multi_week_analyzer.py         # Analysis engine
├── weeks_index.json               # Week metadata
└── [other files...]
```

## Access Instructions for Company:

Once deployed, share this with your company:

**🔗 Dashboard URL**: `https://YOUR-USERNAME.github.io/attendance-analysis/master_dashboard.html`

**📱 Mobile Friendly**: Works on all devices
**🔄 Updates**: Automatic when you push new weeks
**📊 Features**: Interactive charts, week selection, detailed reports

## Troubleshooting:

1. **404 Error**: Make sure branch is "main" and folder is "/ (root)"
2. **Charts Not Loading**: GitHub Pages serves over HTTPS, charts should work
3. **Path Issues**: All paths in HTML are relative, should work automatically
4. **Updates Not Showing**: Changes take 1-2 minutes to deploy

## Security Notes:

- ✅ Student names and personal data are not exposed
- ✅ Only aggregate statistics are shown
- ✅ Excel files with detailed data stay private in repository
- ✅ GitHub repository can be private if company prefers