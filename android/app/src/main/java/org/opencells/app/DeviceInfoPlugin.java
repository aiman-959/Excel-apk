package org.opencells.app;

import android.app.ActivityManager;
import android.content.Context;
import android.os.Build;
import android.view.Display;
import android.view.Window;
import android.view.WindowManager;
import com.getcapacitor.JSArray;
import com.getcapacitor.JSObject;
import com.getcapacitor.Plugin;
import com.getcapacitor.PluginCall;
import com.getcapacitor.PluginMethod;
import com.getcapacitor.annotation.CapacitorPlugin;

/**
 * Reports hardware facts used to choose a performance tier. Works on every
 * Android chip vendor (Qualcomm, MediaTek, Samsung Exynos, Google Tensor, ...):
 * only RAM and core count drive the tier; SoC names are informational.
 */
@CapacitorPlugin(name = "DeviceInfo")
public class DeviceInfoPlugin extends Plugin {

    @PluginMethod
    public void getInfo(PluginCall call) {
        ActivityManager am = (ActivityManager) getContext().getSystemService(Context.ACTIVITY_SERVICE);
        ActivityManager.MemoryInfo mi = new ActivityManager.MemoryInfo();
        am.getMemoryInfo(mi);

        JSObject ret = new JSObject();
        ret.put("totalMem", mi.totalMem);
        ret.put("availMem", mi.availMem);
        ret.put("cores", Runtime.getRuntime().availableProcessors());
        ret.put("isLowRamDevice", am.isLowRamDevice());
        ret.put("sdkInt", Build.VERSION.SDK_INT);
        ret.put("hardware", Build.HARDWARE);
        String socMan = "";
        String socModel = "";
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.S) {
            socMan = Build.SOC_MANUFACTURER;
            socModel = Build.SOC_MODEL;
        }
        ret.put("socManufacturer", socMan);
        ret.put("socModel", socModel);
        JSArray abis = new JSArray();
        for (String abi : Build.SUPPORTED_ABIS) abis.put(abi);
        ret.put("abis", abis);
        ret.put("maxRefreshRate", maxRefreshRate());
        call.resolve(ret);
    }

    /** Pick the display mode with the highest refresh rate at the current resolution (or reset). */
    @PluginMethod
    public void setHighRefreshRate(PluginCall call) {
        final boolean enabled = Boolean.TRUE.equals(call.getBoolean("enabled", false));
        getActivity().runOnUiThread(() -> {
            float applied = 60f;
            try {
                Window w = getActivity().getWindow();
                Display d = display();
                WindowManager.LayoutParams lp = w.getAttributes();
                if (d != null && Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
                    Display.Mode cur = d.getMode();
                    Display.Mode best = cur;
                    if (enabled) {
                        for (Display.Mode m : d.getSupportedModes()) {
                            if (m.getPhysicalWidth() == cur.getPhysicalWidth() && m.getPhysicalHeight() == cur.getPhysicalHeight()
                                    && m.getRefreshRate() > best.getRefreshRate()) {
                                best = m;
                            }
                        }
                        lp.preferredDisplayModeId = best.getModeId();
                        lp.preferredRefreshRate = best.getRefreshRate();
                    } else {
                        lp.preferredDisplayModeId = 0;
                        lp.preferredRefreshRate = 0;
                    }
                    w.setAttributes(lp);
                    applied = enabled ? best.getRefreshRate() : d.getRefreshRate();
                }
            } catch (Exception ignored) {
            }
            JSObject ret = new JSObject();
            ret.put("refreshRate", applied);
            call.resolve(ret);
        });
    }

    private Display display() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) return getActivity().getDisplay();
        WindowManager wm = (WindowManager) getContext().getSystemService(Context.WINDOW_SERVICE);
        return wm.getDefaultDisplay();
    }

    private float maxRefreshRate() {
        Display d = display();
        if (d == null) return 60f;
        float max = d.getRefreshRate();
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M) {
            for (Display.Mode m : d.getSupportedModes()) max = Math.max(max, m.getRefreshRate());
        }
        return max;
    }
}
