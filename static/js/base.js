/* ========================================
   BASE JAVASCRIPT
   - Sidebar toggle
   - User menu dropdown
   - Image zoom modal
   - Menu search
   ======================================== */

document.addEventListener('DOMContentLoaded', function () {
    initSidebar();
    initUserMenu();
    initZoomModal();
    initMenuSearch();
    initHeaderHeight();
});

/**
 * Initialize dynamic header height adjustment
 */
function initHeaderHeight() {
    const root = document.documentElement;
    const header = document.querySelector('.header');

    if (header) {
        const adjustPadding = () => {
            const height = header.offsetHeight;
            root.style.setProperty('--header-height', `${height}px`);
        };

        // Initial adjustment
        adjustPadding();

        // Adjust on resize
        window.addEventListener('resize', adjustPadding);
    }
}

/**
 * Initialize sidebar toggle functionality
 */
function initSidebar() {
    const sidebar = document.getElementById('sidebar');
    const mainContent = document.getElementById('mainContent');
    const sidebarToggle = document.getElementById('sidebarToggle');
    const mobileToggle = document.getElementById('mobileToggle');
    const sidebarOverlay = document.getElementById('sidebarOverlay');

    // Desktop sidebar toggle
    if (sidebarToggle) {
        sidebarToggle.addEventListener('click', function () {
            sidebar.classList.toggle('collapsed');
            mainContent.classList.toggle('sidebar-collapsed');
        });
    }

    // Mobile sidebar toggle
    if (mobileToggle) {
        mobileToggle.addEventListener('click', function () {
            sidebar.classList.toggle('active');
            if (sidebarOverlay) {
                sidebarOverlay.classList.toggle('active');
            }
        });
    }

    // Close sidebar when clicking overlay
    if (sidebarOverlay) {
        sidebarOverlay.addEventListener('click', function () {
            sidebar.classList.remove('active');
            sidebarOverlay.classList.remove('active');
        });
    }

    // Close mobile sidebar when clicking menu item
    const menuItems = document.querySelectorAll('.menu-item');
    menuItems.forEach(item => {
        item.addEventListener('click', function () {
            if (window.innerWidth <= 768) {
                sidebar.classList.remove('active');
                if (sidebarOverlay) {
                    sidebarOverlay.classList.remove('active');
                }
            }
        });
    });
}

/**
 * Initialize user menu dropdown
 */
function initUserMenu() {
    const userMenuButton = document.getElementById('userMenuButton');
    const userDropdown = document.getElementById('userDropdown');

    if (userMenuButton && userDropdown) {
        // Toggle dropdown on button click
        userMenuButton.addEventListener('click', function (e) {
            e.stopPropagation();
            userDropdown.classList.toggle('show');
        });

        // Close dropdown when clicking outside
        document.addEventListener('click', function (e) {
            if (!userMenuButton.contains(e.target) && !userDropdown.contains(e.target)) {
                userDropdown.classList.remove('show');
            }
        });
    }
}

/**
 * Initialize image zoom modal
 */
function initZoomModal() {
    const zoomModal = document.getElementById('zoomModal');
    const zoomImage = document.getElementById('zoomImage');
    const zoomClose = document.getElementById('zoomClose');

    // Add click event to all chart images
    document.addEventListener('click', function (e) {
        if (e.target.tagName === 'IMG' && e.target.closest('.chart-card')) {
            zoomModal.classList.add('active');
            zoomImage.src = e.target.src;
        }
    });

    // Close modal on close button click
    if (zoomClose) {
        zoomClose.addEventListener('click', function () {
            zoomModal.classList.remove('active');
        });
    }

    // Close modal on background click
    if (zoomModal) {
        zoomModal.addEventListener('click', function (e) {
            if (e.target === zoomModal) {
                zoomModal.classList.remove('active');
            }
        });
    }

    // Close modal on ESC key
    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape' && zoomModal.classList.contains('active')) {
            zoomModal.classList.remove('active');
        }
    });
}

/**
 * Initialize menu search functionality
 */
function initMenuSearch() {
    const searchBox = document.getElementById('searchBox');
    const menuItems = document.querySelectorAll('.menu-item');

    if (searchBox) {
        searchBox.addEventListener('input', debounce(function (e) {
            const searchTerm = e.target.value.toLowerCase();

            menuItems.forEach(item => {
                const text = item.textContent.toLowerCase();
                if (text.includes(searchTerm)) {
                    item.style.display = 'flex';
                } else {
                    item.style.display = 'none';
                }
            });
        }, 300));
    }
}

/**
 * Render tabs for data display
 */
function renderTabs(containerId, tabs, onTabClick) {
    const container = document.getElementById(containerId);
    if (!container) return;

    let html = '<div class="excel-tabs-container"><div class="excel-tabs">';

    tabs.forEach((tab, index) => {
        const activeClass = index === 0 ? 'active' : '';
        html += `
            <button class="excel-tab ${activeClass}" data-tab="${tab.id}">
                ${tab.label}
            </button>
        `;
    });

    html += '</div></div>';
    container.innerHTML = html;

    // Add click event listeners
    const tabButtons = container.querySelectorAll('.excel-tab');
    tabButtons.forEach(btn => {
        btn.addEventListener('click', function () {
            // Remove active class from all tabs
            tabButtons.forEach(b => b.classList.remove('active'));

            // Add active class to clicked tab
            this.classList.add('active');

            // Call the callback function
            if (onTabClick) {
                onTabClick(this.dataset.tab);
            }
        });
    });
}

/**
 * Render image tabs
 */
function renderImageTabs(containerId, tabs, onTabClick) {
    const container = document.getElementById(containerId);
    if (!container) return;

    let html = '<div class="image-tabs-container"><div class="image-tabs">';

    tabs.forEach((tab, index) => {
        const activeClass = index === 0 ? 'active' : '';
        html += `
            <button class="image-tab ${activeClass}" data-chart="${tab.id}">
                ${tab.label}
            </button>
        `;
    });

    html += '</div></div>';

    // Insert before the chart display area
    const existingTabs = container.querySelector('.image-tabs-container');
    if (existingTabs) {
        existingTabs.remove();
    }

    container.insertAdjacentHTML('afterbegin', html);

    // Add click event listeners
    const tabButtons = container.querySelectorAll('.image-tab');
    tabButtons.forEach(btn => {
        btn.addEventListener('click', function () {
            // Remove active class from all tabs
            tabButtons.forEach(b => b.classList.remove('active'));

            // Add active class to clicked tab
            this.classList.add('active');

            // Call the callback function
            if (onTabClick) {
                onTabClick(this.dataset.chart);
            }
        });
    });
}

/**
 * Display chart image
 */
function displayChart(containerId, imagePath, title = '', subtitle = '') {
    const container = document.getElementById(containerId);
    if (!container) return;

    const html = `
        <div class="chart-card">
            ${title ? `
                <div class="chart-card-header">
                    <div class="chart-card-title">${title}</div>
                    ${subtitle ? `<div class="chart-card-subtitle">${subtitle}</div>` : ''}
                </div>
            ` : ''}
            <div class="chart-card-body">
                <img src="${imagePath}" alt="${title}" loading="lazy">
            </div>
        </div>
    `;

    container.innerHTML = html;
}

// Export functions for use in page-specific scripts
if (typeof window !== 'undefined') {
    window.renderTabs = renderTabs;
    window.renderImageTabs = renderImageTabs;
    window.displayChart = displayChart;
}
