# -*- coding: utf-8 -*-
with open('app_kiem_diem_chi_bo_2.html', 'r', encoding='utf-8') as f:
    c = f.read()

# 1. Add buttons in Tab 4 leader-assign-tools
old_leader_tools = """            <!-- Phân quyền điều phối: Chỉ Bí thư / Chi ủy được thấy nút phân công -->
            <div id="leader-assign-tools" class="flex items-center gap-2">
              <button onclick="openManualAssignModal()"
                class="btn-3d text-xs bg-vtb-navy hover:bg-vtb-navyDark text-white px-3.5 py-1.5 rounded-xl font-bold flex items-center gap-1.5 shadow transition">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                    d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                </svg>
                Phân Công Nhận Xét
              </button>

              <button onclick="autoAssignBalancedPeerReviews()"
                class="btn-3d text-xs bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-xl font-bold flex items-center gap-1.5 shadow transition">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                    d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                </svg>
                Tự Động Chia Cân Bằng (2 người/đồng chí)
              </button>
            </div>"""

new_leader_tools = """            <!-- Phân quyền điều phối: Dành riêng cho Bí thư / Ban Chi ủy -->
            <div id="leader-assign-tools" class="flex items-center flex-wrap gap-2">
              <button onclick="openReviewMeetingReportModal()"
                class="btn-3d text-xs bg-gradient-to-r from-red-700 to-rose-800 hover:from-red-800 hover:to-rose-900 text-white px-3.5 py-1.5 rounded-xl font-bold flex items-center gap-1.5 shadow-md transition" title="Kết xuất Báo cáo tổng hợp nhận xét phục vụ Bí thư điều hành cuộc họp">
                <i class="fa-solid fa-file-word text-amber-300"></i>
                <span>Báo Cáo Nhận Xét Điều Hành (.docx)</span>
              </button>

              <button onclick="openAllReviewsModal()"
                class="btn-3d text-xs bg-gradient-to-r from-indigo-700 to-purple-800 hover:from-indigo-800 hover:to-purple-900 text-white px-3.5 py-1.5 rounded-xl font-bold flex items-center gap-1.5 shadow-md transition" title="Bí thư xem toàn bộ nhận xét của các Đảng viên trong Chi bộ">
                <i class="fa-solid fa-layer-group text-amber-300"></i>
                <span>Xem Toàn Bộ Nhận Xét</span>
              </button>

              <button onclick="openManualAssignModal()"
                class="btn-3d text-xs bg-vtb-navy hover:bg-vtb-navyDark text-white px-3.5 py-1.5 rounded-xl font-bold flex items-center gap-1.5 shadow transition">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                    d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z" />
                </svg>
                Phân Công Nhận Xét
              </button>

              <button onclick="autoAssignBalancedPeerReviews()"
                class="btn-3d text-xs bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-xl font-bold flex items-center gap-1.5 shadow transition">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                    d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15" />
                </svg>
                Tự Động Chia Cân Bằng (2 người/đồng chí)
              </button>
            </div>"""

assert old_leader_tools in c, "old_leader_tools not found"
c = c.replace(old_leader_tools, new_leader_tools)
print("1. Added Secretary buttons to leader-assign-tools")

# 2. Update renderMyReviewTasks to support Secretary reviewing all members
old_render_tasks_split = """      // Phân tách: Cán bộ ĐƯỢC PHÂN CÔNG vs Cán bộ KHÁC (BỊ KHÓA)
      const targetMems = targets.map(tid => appData.members.find(x => x.id === tid)).filter(Boolean);
      const lockedMembers = appData.members.filter(m => m.id !== currentReviewer.id && !targets.includes(m.id));
      const targetNamesStr = targetMems.map(m => m.name).join(', ');"""

new_render_tasks_split = """      const isLeader = (currentReviewer.role_party === 'Bí thư Chi bộ' || currentReviewer.role_party === 'Phó Bí thư Chi bộ');
      // Phân tách: Cán bộ ĐƯỢC PHÂN CÔNG vs Cán bộ KHÁC
      const targetMems = targets.map(tid => appData.members.find(x => x.id === tid)).filter(Boolean);
      const otherMembers = appData.members.filter(m => m.id !== currentReviewer.id && !targets.includes(m.id));
      const targetNamesStr = targetMems.map(m => m.name).join(', ');"""

assert old_render_tasks_split in c, "old_render_tasks_split not found"
c = c.replace(old_render_tasks_split, new_render_tasks_split)

old_banner = """      // 1. BANNER PHÂN QUYỀN RÕ RÀNG
      const permissionBanner = document.createElement('div');
      permissionBanner.className = 'col-span-1 md:col-span-2 p-4 bg-gradient-to-r from-emerald-50 via-teal-50 to-blue-50 border border-emerald-200 rounded-2xl flex flex-wrap items-center justify-between gap-3 shadow-xs mb-1';
      permissionBanner.innerHTML = `
        <div class="flex items-start gap-3">
          <span class="w-9 h-9 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 text-white flex items-center justify-center text-sm font-bold shadow-sm shrink-0 mt-0.5">
            <i class="fa-solid fa-user-shield"></i>
          </span>
          <div>
            <div class="text-xs font-black text-slate-800 flex flex-wrap items-center gap-2">
              <span>Phân quyền nhận xét:</span>
              <span class="text-vtb-red uppercase font-black text-sm">${currentReviewer.name}</span>
              <span class="px-2 py-0.5 rounded text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">
                ${currentReviewer.role_party}
              </span>
            </div>
            <p class="text-xs text-slate-700 mt-1">
              🎯 Đồng chí <strong>chỉ được phép nhận xét</strong> cho <strong>${targetMems.length} cán bộ</strong> đã được Chi bộ phân công: 
              <span class="font-extrabold text-vtb-navy">${targetNamesStr || 'Chưa phân công'}</span>.
            </p>
            <p class="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
              <i class="fa-solid fa-lock text-amber-600"></i>
              <span>Toàn bộ <strong>${lockedMembers.length} cán bộ còn lại</strong> trong Chi bộ đã được <strong>khóa chức năng nhận xét</strong> theo quy chế sinh hoạt Đảng.</span>
            </p>
          </div>
        </div>

        <div class="flex items-center gap-2">
          <span class="text-xs font-black px-3 py-1.5 bg-white text-emerald-800 rounded-xl border border-emerald-300 shadow-xs flex items-center gap-1.5">
            <i class="fa-solid fa-circle-check text-emerald-600"></i> Được phân công (${targetMems.length})
          </span>
          <span class="text-xs font-black px-3 py-1.5 bg-white text-slate-600 rounded-xl border border-slate-300 shadow-xs flex items-center gap-1.5">
            <i class="fa-solid fa-lock text-amber-600"></i> Đã khóa nhận xét (${lockedMembers.length})
          </span>
        </div>
      `;
      container.appendChild(permissionBanner);"""

new_banner = """      // 1. BANNER PHÂN QUYỀN RÕ RÀNG
      const permissionBanner = document.createElement('div');
      permissionBanner.className = 'col-span-1 md:col-span-2 p-4 bg-gradient-to-r from-emerald-50 via-teal-50 to-blue-50 border border-emerald-200 rounded-2xl flex flex-wrap items-center justify-between gap-3 shadow-xs mb-1';
      
      if (isLeader) {
        permissionBanner.innerHTML = `
          <div class="flex items-start gap-3">
            <span class="w-10 h-10 rounded-xl bg-gradient-to-br from-amber-500 to-red-600 text-white flex items-center justify-center text-base font-bold shadow-md shrink-0 mt-0.5">
              👑
            </span>
            <div>
              <div class="text-xs font-black text-slate-800 flex flex-wrap items-center gap-2">
                <span>Quyền hạn điều hành của:</span>
                <span class="text-vtb-red uppercase font-black text-sm">${currentReviewer.name}</span>
                <span class="px-2.5 py-0.5 rounded-full text-[10px] font-black bg-amber-400 text-slate-950 border border-amber-500 shadow-xs">
                  ${currentReviewer.role_party}
                </span>
              </div>
              <p class="text-xs text-slate-800 mt-1">
                ⭐ <strong>Đặc quyền Bí thư Chi bộ:</strong> Được <strong>xem toàn bộ nhận xét</strong> của tất cả Đảng viên và <strong>có quyền nhận xét, đánh giá kết luận toàn bộ 11 Đảng viên còn lại</strong> để phục vụ điều hành cuộc họp kiểm điểm.
              </p>
              <div class="flex flex-wrap items-center gap-2 mt-2">
                <button type="button" onclick="openReviewMeetingReportModal()" class="px-3 py-1 bg-red-700 hover:bg-red-800 text-white rounded-lg text-xs font-bold shadow-xs transition flex items-center gap-1.5">
                  <i class="fa-solid fa-file-word text-amber-300"></i>
                  <span>Xuất Báo Cáo Nhận Xét Điều Hành (.docx)</span>
                </button>
                <button type="button" onclick="openAllReviewsModal()" class="px-3 py-1 bg-indigo-700 hover:bg-indigo-800 text-white rounded-lg text-xs font-bold shadow-xs transition flex items-center gap-1.5">
                  <i class="fa-solid fa-list-check text-amber-300"></i>
                  <span>Bảng Tổng Hợp Nhận Xét Toàn Chi Bộ</span>
                </button>
              </div>
            </div>
          </div>

          <div class="flex items-center gap-2 shrink-0">
            <span class="text-xs font-black px-3 py-1.5 bg-white text-emerald-800 rounded-xl border border-emerald-300 shadow-xs flex items-center gap-1.5">
              <i class="fa-solid fa-check-double text-emerald-600"></i> Nhận xét tất cả (11 Đ/c)
            </span>
          </div>
        `;
      } else {
        permissionBanner.innerHTML = `
          <div class="flex items-start gap-3">
            <span class="w-9 h-9 rounded-xl bg-gradient-to-br from-emerald-600 to-teal-700 text-white flex items-center justify-center text-sm font-bold shadow-sm shrink-0 mt-0.5">
              <i class="fa-solid fa-user-shield"></i>
            </span>
            <div>
              <div class="text-xs font-black text-slate-800 flex flex-wrap items-center gap-2">
                <span>Phân quyền nhận xét:</span>
                <span class="text-vtb-red uppercase font-black text-sm">${currentReviewer.name}</span>
                <span class="px-2 py-0.5 rounded text-[10px] font-black bg-emerald-100 text-emerald-800 border border-emerald-300">
                  ${currentReviewer.role_party}
                </span>
              </div>
              <p class="text-xs text-slate-700 mt-1">
                🎯 Đồng chí <strong>chỉ được phép nhận xét</strong> cho <strong>${targetMems.length} cán bộ</strong> đã được Chi bộ phân công: 
                <span class="font-extrabold text-vtb-navy">${targetNamesStr || 'Chưa phân công'}</span>.
              </p>
              <p class="text-[11px] text-slate-500 mt-0.5 flex items-center gap-1">
                <i class="fa-solid fa-lock text-amber-600"></i>
                <span>Toàn bộ <strong>${otherMembers.length} cán bộ còn lại</strong> trong Chi bộ đã được <strong>khóa chức năng nhận xét</strong> theo quy chế sinh hoạt Đảng.</span>
              </p>
            </div>
          </div>

          <div class="flex items-center gap-2">
            <span class="text-xs font-black px-3 py-1.5 bg-white text-emerald-800 rounded-xl border border-emerald-300 shadow-xs flex items-center gap-1.5">
              <i class="fa-solid fa-circle-check text-emerald-600"></i> Được phân công (${targetMems.length})
            </span>
            <span class="text-xs font-black px-3 py-1.5 bg-white text-slate-600 rounded-xl border border-slate-300 shadow-xs flex items-center gap-1.5">
              <i class="fa-solid fa-lock text-amber-600"></i> Đã khóa nhận xét (${otherMembers.length})
            </span>
          </div>
        `;
      }
      container.appendChild(permissionBanner);"""

assert old_banner in c, "old_banner not found"
c = c.replace(old_banner, new_banner)

# 3. Update the bottom section of renderMyReviewTasks:
# If isLeader: show other members with the ability to review them!
# If normal member: show locked members.
old_bottom_section = """      // 2. KHU VỰC CÁC CÁN BỘ KHÁC: KHÓA CHỨC NĂNG NHẬN XÉT
      if (lockedMembers.length > 0) {
        const lockedSection = document.createElement('div');
        lockedSection.className = 'col-span-1 md:col-span-2 mt-6 pt-5 border-t border-slate-200 space-y-3';
        lockedSection.innerHTML = `
          <div class="flex flex-wrap items-center justify-between gap-2">
            <h4 class="text-xs font-bold text-slate-700 flex items-center gap-2">
              <i class="fa-solid fa-lock text-amber-600"></i>
              <span>Danh Sách Cán Bộ Khác Không Thuộc Phân Công (${lockedMembers.length} đồng chí - ĐÃ KHÓA NHẬN XÉT)</span>
            </h4>
            <span class="text-[11px] text-slate-400 italic">
              Theo quy định của Chi ủy, đồng chí chỉ nhận xét đúng các cán bộ được phân công
            </span>
          </div>

          <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5">
            ${lockedMembers.map(m => `
              <div class="p-3 bg-slate-50/90 border border-slate-200 rounded-xl flex items-center justify-between gap-2 text-xs opacity-85 hover:opacity-100 transition shadow-xs">
                <div class="min-w-0">
                  <div class="font-bold text-slate-800 truncate flex items-center gap-1.5">
                    <i class="fa-solid fa-user text-slate-400 text-[10px]"></i>
                    <span>Đ/c ${m.name}</span>
                  </div>
                  <div class="text-[10px] text-slate-500 truncate">${m.role_party || 'Đảng viên'} • ${m.title_pro || 'Cán bộ'}</div>
                </div>
                <button type="button" onclick="notifyLockedReview('${m.name}')" class="shrink-0 px-2 py-1 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-[10px] font-bold border border-slate-300 flex items-center gap-1 cursor-pointer transition" title="Bấm để xem thông báo khóa">
                  <i class="fa-solid fa-lock text-amber-700"></i>
                  <span>Khóa nhận xét</span>
                </button>
              </div>
            `).join('')}
          </div>
        `;
        container.appendChild(lockedSection);
      }"""

new_bottom_section = """      // 2. KHU VỰC CÁC CÁN BỘ KHÁC:
      // - Nếu là BÍ THƯ: Mở quyền nhận xét toàn bộ các Đảng viên còn lại!
      // - Nếu là ĐẢNG VIÊN THƯỜNG: Khóa hoàn toàn chức năng nhận xét.
      if (otherMembers.length > 0) {
        const otherSection = document.createElement('div');
        otherSection.className = 'col-span-1 md:col-span-2 mt-6 pt-5 border-t border-slate-200 space-y-4';
        
        if (isLeader) {
          otherSection.innerHTML = `
            <div class="flex flex-wrap items-center justify-between gap-2 bg-amber-50/80 border border-amber-200 p-3.5 rounded-xl">
              <div>
                <h4 class="text-xs font-black text-amber-950 flex items-center gap-2">
                  <i class="fa-solid fa-crown text-amber-600"></i>
                  <span>QUYỀN HẠN BÍ THƯ: NHẬN XÉT & ĐÁNH GIÁ KẾT LUẬN ${otherMembers.length} ĐẢNG VIÊN CÒN LẠI</span>
                </h4>
                <p class="text-[11px] text-amber-800 mt-0.5">
                  Đồng chí Bí thư có thể trực tiếp bổ sung ý kiến chỉ đạo, nhận xét, đánh giá kết luận cho từng Đảng viên dưới đây.
                </p>
              </div>
              <span class="text-xs font-bold px-2.5 py-1 bg-amber-200 text-amber-900 rounded-lg">
                Được phép nhận xét 100%
              </span>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              ${otherMembers.map(m => {
                const revList = appData.reviews[m.id] || [];
                let existingRev = revList.find(r => r.reviewer_id === currentReviewer.id);
                let goodVal = existingRev ? (existingRev.good || '') : `Đồng chí ${m.name} luôn giữ vững phẩm chất chính trị, chấp hành tốt kỷ luật nghiệp vụ và hoàn thành tốt chức trách nhiệm vụ.`;
                let limitVal = existingRev ? (existingRev.limit || '') : 'Cần chủ động hơn trong công tác nghiên cứu học tập nghị quyết và bám sát tiến độ chuyên môn.';
                let solVal = existingRev ? (existingRev.solution || '') : 'Phát huy kết quả đạt được, tăng cường kỷ cương và chủ động phối hợp với các đồng chí trong Chi bộ.';

                // Tìm xem ai khác đã nhận xét đồng chí này
                const otherRevs = revList.filter(r => r.reviewer_id !== currentReviewer.id);
                const otherRevNames = otherRevs.map(r => r.reviewer_name).join(', ') || 'Chưa có ý kiến nhận xét chéo';

                return `
                  <div class="border-2 border-amber-200/90 hover:border-amber-400 rounded-2xl p-4 bg-gradient-to-br from-white to-amber-50/20 shadow-sm space-y-3.5 transition">
                    <div class="flex flex-wrap items-center justify-between border-b pb-2.5 gap-2">
                      <div>
                        <span class="text-[10px] uppercase font-black px-2.5 py-0.5 rounded-full bg-amber-100 text-amber-900 border border-amber-300 inline-block mb-1">
                          <i class="fa-solid fa-crown mr-1"></i> BÍ THƯ ĐÁNH GIÁ KẾT LUẬN
                        </span>
                        <h4 class="font-bold text-sm text-vtb-navy flex items-center gap-1.5">
                          <span>Đồng chí:</span> <span class="text-base text-red-700 font-black">${m.name}</span>
                        </h4>
                        <p class="text-[11px] text-slate-500 font-medium">${m.role_party || 'Đảng viên'} • ${m.title_pro || 'Cán bộ'}</p>
                      </div>
                      <div class="text-right">
                        <span class="text-xs font-bold text-emerald-800 bg-emerald-100 border border-emerald-200 px-2 py-0.5 rounded-lg inline-block">
                          Tự chấm: ${m.self_score}đ
                        </span>
                        <div class="text-[10px] text-slate-500 mt-1 italic">Được góp ý bởi: <strong>${otherRevNames}</strong></div>
                      </div>
                    </div>

                    <div class="space-y-3 text-xs">
                      <div class="space-y-1">
                        <label class="font-bold text-slate-700 flex items-center gap-1">
                          <i class="fa-solid fa-circle-check text-emerald-600"></i>
                          <span>1. Ưu điểm nổi bật / Nhất trí:</span>
                        </label>
                        <textarea id="rev-good-${m.id}" rows="3" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-1 focus:ring-amber-500 text-xs leading-relaxed font-normal whitespace-pre-wrap">${goodVal}</textarea>
                      </div>

                      <div class="space-y-1">
                        <label class="font-bold text-slate-700 flex items-center gap-1">
                          <i class="fa-solid fa-triangle-exclamation text-amber-600"></i>
                          <span>2. Hạn chế, khuyết điểm cần khắc phục:</span>
                        </label>
                        <textarea id="rev-limit-${m.id}" rows="3" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-1 focus:ring-amber-500 text-xs leading-relaxed font-normal whitespace-pre-wrap">${limitVal}</textarea>
                      </div>

                      <div class="space-y-1">
                        <label class="font-bold text-slate-700 flex items-center gap-1">
                          <i class="fa-solid fa-lightbulb text-blue-600"></i>
                          <span>3. Đề nghị, kết luận của Bí thư Chi bộ:</span>
                        </label>
                        <textarea id="rev-sol-${m.id}" rows="3" class="w-full p-2.5 border border-slate-300 rounded-xl focus:ring-1 focus:ring-amber-500 text-xs leading-relaxed font-normal whitespace-pre-wrap">${solVal}</textarea>
                      </div>
                    </div>

                    <div class="flex items-center justify-between pt-2 border-t border-slate-100">
                      <button type="button" onclick="applyFullPeerSuggestion(${m.id})" class="text-[11px] px-2.5 py-1 bg-purple-50 hover:bg-purple-100 text-purple-800 font-bold rounded-lg border border-purple-200 transition">
                        <i class="fa-solid fa-wand-magic-sparkles text-purple-600 mr-1"></i> Nạp mẫu gợi ý
                      </button>
                      <button onclick="savePeerReview(${m.id})" class="btn-3d text-xs bg-amber-600 hover:bg-amber-700 text-white font-bold px-4 py-2 rounded-xl shadow-md transition flex items-center gap-1.5">
                        <i class="fa-solid fa-floppy-disk"></i>
                        <span>Lưu Đánh Giá Của Bí Thư</span>
                      </button>
                    </div>
                  </div>
                `;
              }).join('')}
            </div>
          `;
        } else {
          otherSection.innerHTML = `
            <div class="flex flex-wrap items-center justify-between gap-2">
              <h4 class="text-xs font-bold text-slate-700 flex items-center gap-2">
                <i class="fa-solid fa-lock text-amber-600"></i>
                <span>Danh Sách Cán Bộ Khác Không Thuộc Phân Công (${otherMembers.length} đồng chí - ĐÃ KHÓA NHẬN XÉT)</span>
              </h4>
              <span class="text-[11px] text-slate-400 italic">
                Theo quy định của Chi ủy, đồng chí chỉ nhận xét đúng các cán bộ được phân công
              </span>
            </div>

            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-2.5">
              ${otherMembers.map(m => `
                <div class="p-3 bg-slate-50/90 border border-slate-200 rounded-xl flex items-center justify-between gap-2 text-xs opacity-85 hover:opacity-100 transition shadow-xs">
                  <div class="min-w-0">
                    <div class="font-bold text-slate-800 truncate flex items-center gap-1.5">
                      <i class="fa-solid fa-user text-slate-400 text-[10px]"></i>
                      <span>Đ/c ${m.name}</span>
                    </div>
                    <div class="text-[10px] text-slate-500 truncate">${m.role_party || 'Đảng viên'} • ${m.title_pro || 'Cán bộ'}</div>
                  </div>
                  <button type="button" onclick="notifyLockedReview('${m.name}')" class="shrink-0 px-2 py-1 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg text-[10px] font-bold border border-slate-300 flex items-center gap-1 cursor-pointer transition" title="Bấm để xem thông báo khóa">
                    <i class="fa-solid fa-lock text-amber-700"></i>
                    <span>Khóa nhận xét</span>
                  </button>
                </div>
              `).join('')}
            </div>
          `;
        }
        container.appendChild(otherSection);
      }"""

assert old_bottom_section in c, "old_bottom_section not found"
c = c.replace(old_bottom_section, new_bottom_section)
print("2. Updated renderMyReviewTasks with Secretary full-review capability")

# 4. Update savePeerReview permissions
old_save_guard = """      // KIỂM TRA PHÂN QUYỀN CHẶT CHẼ: Chặn hoàn toàn nếu không được phân công
      if (!assignedTargets.includes(tid)) {
        showToast('Đồng chí không được phân công nhận xét cán bộ này! Chức năng đã bị khóa.', 'error');
        alert(`⚠️ Quyền hạn không hợp lệ!\\n\\nĐồng chí ${currentReviewer.name} chỉ được phân công nhận xét cho:\\n${assignedTargets.map(id => {
          const m = appData.members.find(x => x.id === id);
          return '• Đ/c ' + (m ? m.name : id);
        }).join('\\n')}\\n\\nChức năng nhận xét các cán bộ khác đã bị khóa theo quy chế sinh hoạt Chi bộ.`);
        return;
      }"""

new_save_guard = """      const isLeader = (currentReviewer.role_party === 'Bí thư Chi bộ' || currentReviewer.role_party === 'Phó Bí thư Chi bộ');

      // KIỂM TRA PHÂN QUYỀN CHẶT CHẼ:
      // Đảng viên thường bị chặn nếu không thuộc phân công.
      // Riêng Bí thư / Ban Chi ủy được toàn quyền nhận xét tất cả Đảng viên để phục vụ điều hành cuộc họp!
      if (!isLeader && !assignedTargets.includes(tid)) {
        showToast('Đồng chí không được phân công nhận xét cán bộ này! Chức năng đã bị khóa.', 'error');
        alert(`⚠️ Quyền hạn không hợp lệ!\\n\\nĐồng chí ${currentReviewer.name} chỉ được phân công nhận xét cho:\\n${assignedTargets.map(id => {
          const m = appData.members.find(x => x.id === id);
          return '• Đ/c ' + (m ? m.name : id);
        }).join('\\n')}\\n\\nChức năng nhận xét các cán bộ khác đã bị khóa theo quy chế sinh hoạt Chi bộ.`);
        return;
      }"""

assert old_save_guard in c, "old_save_guard not found"
c = c.replace(old_save_guard, new_save_guard)

# Update helper functions guards:
c = c.replace(
    "if (!assignedTargets.includes(tid)) {\n        showToast('Chức năng nhận xét cán bộ này đã bị khóa!', 'warning');\n        return;\n      }",
    "const isLeader = (currentReviewer.role_party === 'Bí thư Chi bộ' || currentReviewer.role_party === 'Phó Bí thư Chi bộ');\n      if (!isLeader && !assignedTargets.includes(tid)) {\n        showToast('Chức năng nhận xét cán bộ này đã bị khóa!', 'warning');\n        return;\n      }"
)
print("3. Updated savePeerReview guards for Secretary")

with open('app_kiem_diem_chi_bo_2.html', 'w', encoding='utf-8') as f:
    f.write(c)

print("Saved stage 1 of Secretary features successfully!")
