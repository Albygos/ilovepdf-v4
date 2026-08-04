/**
 * Advanced Google AdSense Manager for ilovespdfs.in
 * Automatically injects AdSense slots and auto-ads across all homepages, tool pages, and SEO pages.
 */

document.addEventListener("DOMContentLoaded", function () {
    const basePath = '/ads/';

    async function loadAdSnippet(filename) {
        try {
            const response = await fetch(basePath + filename);
            if (response.ok) {
                return await response.text();
            }
        } catch (e) {
            console.warn("Could not load ad snippet: " + filename);
        }
        return null;
    }

    function insertAd(html, targetSelector, position = 'afterend', fallbackSelector = null) {
        if (!html || !html.trim()) return;
        
        // Wrap ad in a responsive zero-CLS container
        const adContainer = document.createElement('div');
        adContainer.className = 'dynamic-ad-container';
        adContainer.style.cssText = 'margin: 24px auto; max-width: 1200px; text-align: center; overflow: hidden; min-height: 90px; z-index: 100; position: relative;';
        adContainer.innerHTML = html;

        let target = document.querySelector(targetSelector);
        if (!target && fallbackSelector) {
            target = document.querySelector(fallbackSelector);
        }

        if (target && target.parentNode) {
            if (position === 'afterend') {
                target.parentNode.insertBefore(adContainer, target.nextSibling);
            } else if (position === 'beforebegin') {
                target.parentNode.insertBefore(adContainer, target);
            } else if (position === 'beforeend') {
                target.appendChild(adContainer);
            }
        }
    }

    async function initializeAds() {
        // 1. Inject Auto Ads script into <head>
        const autoAdsHtml = await loadAdSnippet('auto-ads.html');
        if (autoAdsHtml && autoAdsHtml.includes('<script')) {
            const tempDiv = document.createElement('div');
            tempDiv.innerHTML = autoAdsHtml;
            tempDiv.querySelectorAll('script').forEach(oldScript => {
                const newScript = document.createElement('script');
                Array.from(oldScript.attributes).forEach(attr => newScript.setAttribute(attr.name, attr.value));
                if (oldScript.innerHTML) newScript.textContent = oldScript.innerHTML;
                document.head.appendChild(newScript);
            });
        }

        // 2. Load Top Banner (Above fold / below navbar/hero)
        const topBanner = await loadAdSnippet('top-banner.html');
        if (topBanner) {
            insertAd(topBanner, '.hero-section', 'afterend', 'header');
        }

        // 3. Load Bottom Banner (Above footer / below workspace)
        const bottomBanner = await loadAdSnippet('bottom-banner.html');
        if (bottomBanner) {
            insertAd(bottomBanner, '.tools-grid-wrapper', 'afterend', '.tool-workspace');
        }

        // 4. Load In-Article Ad (Inside SEO content paragraphs)
        const inArticle = await loadAdSnippet('in-article.html');
        if (inArticle) {
            const seoContainer = document.querySelector('.seo-content-section') || document.querySelector('.seo-content') || document.querySelector('main');
            if (seoContainer) {
                const firstP = seoContainer.querySelector('p');
                if (firstP) {
                    insertAd(inArticle, 'p', 'afterend');
                }
            }
        }

        // Push AdSense ads after injection
        setTimeout(() => {
            document.querySelectorAll('.dynamic-ad-container script').forEach(script => {
                if (script.textContent.includes('adsbygoogle') || script.innerHTML.includes('adsbygoogle')) {
                    try {
                        (window.adsbygoogle = window.adsbygoogle || []).push({});
                    } catch (e) {}
                }
            });
        }, 300);
    }

    initializeAds();
});
