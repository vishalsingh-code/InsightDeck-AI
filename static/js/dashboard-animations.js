// Enhanced Dashboard Animations and Interactions

document.addEventListener('DOMContentLoaded', function() {
    // Add smooth animations to feature cards
    const observerOptions = {
        threshold: 0.1,
        rootMargin: '0px 0px -50px 0px'
    };

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.style.animation = 'slideInUp 0.6s ease-out forwards';
            }
        });
    }, observerOptions);

    // Observe feature cards
    document.querySelectorAll('.feature-list li, .stat-card').forEach(card => {
        card.style.opacity = '0';
        card.style.transform = 'translateY(30px)';
        observer.observe(card);
    });

    // Add CSS for animations
    const style = document.createElement('style');
    style.textContent = `
        @keyframes slideInUp {
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        
        .feature-list li {
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .stat-card {
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .upload-area {
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
        }
        
        .btn {
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }
    `;
    document.head.appendChild(style);

    // Enhanced file upload feedback
    const fileInput = document.getElementById('file');
    const uploadArea = document.querySelector('.upload-area');
    
    if (fileInput && uploadArea) {
        fileInput.addEventListener('change', function() {
            if (this.files && this.files[0]) {
                const fileName = this.files[0].name;
                const fileSize = (this.files[0].size / 1024 / 1024).toFixed(2);
                
                uploadArea.querySelector('.upload-text').innerHTML = `
                    <div style="color: #10b981; font-weight: 600;">
                        ✅ File Selected: ${fileName}
                    </div>
                    <div style="color: #6b7280; font-size: 0.9em; margin-top: 8px;">
                        Size: ${fileSize} MB
                    </div>
                `;
                
                uploadArea.style.borderColor = '#10b981';
                uploadArea.style.background = 'linear-gradient(145deg, #ecfdf5 0%, #d1fae5 100%)';
            }
        });
    }

    // Add ripple effect to buttons
    document.querySelectorAll('.btn').forEach(button => {
        button.addEventListener('click', function(e) {
            const ripple = document.createElement('span');
            const rect = this.getBoundingClientRect();
            const size = Math.max(rect.width, rect.height);
            const x = e.clientX - rect.left - size / 2;
            const y = e.clientY - rect.top - size / 2;
            
            ripple.style.cssText = `
                position: absolute;
                width: ${size}px;
                height: ${size}px;
                left: ${x}px;
                top: ${y}px;
                background: rgba(255, 255, 255, 0.3);
                border-radius: 50%;
                transform: scale(0);
                animation: ripple 0.6s ease-out;
                pointer-events: none;
            `;
            
            this.appendChild(ripple);
            
            setTimeout(() => ripple.remove(), 600);
        });
        
        // Add ripple animation
        if (!document.querySelector('#ripple-styles')) {
            const rippleStyle = document.createElement('style');
            rippleStyle.id = 'ripple-styles';
            rippleStyle.textContent = `
                @keyframes ripple {
                    to {
                        transform: scale(2);
                        opacity: 0;
                    }
                }
            `;
            document.head.appendChild(rippleStyle);
        }
    });

    // Smooth scroll for any internal links
    document.querySelectorAll('a[href^="#"]').forEach(anchor => {
        anchor.addEventListener('click', function (e) {
            e.preventDefault();
            const target = document.querySelector(this.getAttribute('href'));
            if (target) {
                target.scrollIntoView({
                    behavior: 'smooth',
                    block: 'start'
                });
            }
        });
    });

    // Add loading states for form submissions
    document.querySelectorAll('form').forEach(form => {
        form.addEventListener('submit', function() {
            const submitBtn = this.querySelector('button[type="submit"]');
            if (submitBtn && !submitBtn.disabled) {
                const originalText = submitBtn.textContent;
                submitBtn.textContent = '⏳ Processing...';
                submitBtn.disabled = true;
                
                // Re-enable after a timeout as fallback
                setTimeout(() => {
                    submitBtn.textContent = originalText;
                    submitBtn.disabled = false;
                }, 30000);
            }
        });
    });
});
